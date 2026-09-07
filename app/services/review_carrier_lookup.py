"""Tra bên vận tải chuẩn từ workbook Hàng ngày cho màn hình review JSON."""

from __future__ import annotations

import logging
from collections import defaultdict
from collections.abc import Sequence
from datetime import date
from pathlib import Path
from typing import Any

from app.services.validation_service import normalize_container

from .carrier_policy import carrier_is_managed_by_daily_sync
from .excel.carrier import carrier_group_for_fee, unique_carriers
from .excel.daily_sync import SOURCE_HEADER_ALIASES
from .excel.headers import HeaderResolutionError, HeaderResolver
from .excel.resolvers import MonthSheetService
from .excel.workbook import WorkbookGateway, ensure_supported_workbook


LOGGER = logging.getLogger(__name__)

_LOOKUP_HEADER_ALIASES = {
    "container": SOURCE_HEADER_ALIASES["container"],
    "carrier_sea": SOURCE_HEADER_ALIASES["sea_transport"],
    "carrier_road": SOURCE_HEADER_ALIASES["transport"],
}


def _value(source: Any, *names: str, default: Any = None) -> Any:
    if isinstance(source, dict):
        for name in names:
            if name in source:
                return source[name]
        return default
    for name in names:
        if hasattr(source, name):
            return getattr(source, name)
    return default


def _invoice_month(value: Any) -> int | None:
    if isinstance(value, date):
        return value.month
    if not isinstance(value, str):
        return None
    try:
        return date.fromisoformat(value).month
    except ValueError:
        return None


def _normalized_container(value: Any) -> str | None:
    try:
        return normalize_container(value)
    except TypeError:
        return None


def _lookup_months(invoice_month: int | None) -> tuple[int, ...] | None:
    if invoice_month is None:
        return None
    return tuple(
        month
        for month in (invoice_month, invoice_month - 1, invoice_month - 2)
        if month >= 1
    )


class ReviewCarrierLookup:
    """Tìm duy nhất carrier SEA/ROAD hiện có theo container của khoản chi."""

    def __init__(
        self,
        daily_path: str | Path,
        *,
        gateway: WorkbookGateway | None = None,
        headers: HeaderResolver | None = None,
        months: MonthSheetService | None = None,
    ) -> None:
        self.daily_path = Path(daily_path) if str(daily_path).strip() else Path()
        self.gateway = gateway or WorkbookGateway()
        self.headers = headers or HeaderResolver()
        self.months = months or MonthSheetService()

    def resolve(self, rows: Sequence[Any]) -> dict[int, str]:
        """Trả carrier theo chỉ số dòng; không bao giờ dùng carrier đang có trong JSON."""

        pending: list[tuple[int, str, str, int | None]] = []
        for index, row in enumerate(rows):
            fee = _value(row, "fee")
            if not carrier_is_managed_by_daily_sync(fee):
                continue
            if _value(row, "carrier") not in (None, ""):
                # Giá trị đã lưu sau review là quyết định của user, không tra đè lại.
                continue
            container = _normalized_container(_value(row, "cont", "container"))
            group = carrier_group_for_fee(str(fee or ""))
            if container is None or group not in {"SEA", "ROAD"}:
                continue
            pending.append(
                (
                    index,
                    container,
                    group,
                    _invoice_month(_value(row, "invoice_date")),
                )
            )
        if not pending:
            return {}

        try:
            source = ensure_supported_workbook(self.daily_path)
        except (TypeError, ValueError):
            return {}
        if not source.is_file():
            LOGGER.warning(
                "Không thể điền bên vận tải khi review vì thiếu file Hàng ngày: %s",
                source,
            )
            return {}

        by_month: dict[tuple[int, str, str], list[str]] = defaultdict(list)
        all_months: dict[tuple[str, str], list[str]] = defaultdict(list)
        # HeaderResolver và phép tra nhiều ô cần truy cập ngẫu nhiên; openpyxl
        # read_only sẽ quét lại XML cho mỗi ``cell()`` và rất chậm trên file thật.
        workbook = self.gateway.load(source, read_only=False, data_only=True)
        try:
            daily_sheets = self.months.daily_sheets(workbook.sheetnames)
            for month, sheet_name in daily_sheets.items():
                worksheet = workbook[sheet_name]
                try:
                    header = self.headers.resolve(
                        worksheet,
                        _LOOKUP_HEADER_ALIASES,
                        required=("container",),
                    )
                except HeaderResolutionError:
                    LOGGER.warning(
                        "Bỏ qua sheet %s khi tra bên vận tải vì không nhận diện được header.",
                        sheet_name,
                    )
                    continue
                carrier_columns = {
                    "SEA": header.columns.get("carrier_sea"),
                    "ROAD": header.columns.get("carrier_road"),
                }
                for row_index in range(
                    header.row_end + 1,
                    int(worksheet.max_row or 0) + 1,
                ):
                    container = _normalized_container(
                        worksheet.cell(row_index, header.columns["container"]).value
                    )
                    if container is None:
                        continue
                    for group, column in carrier_columns.items():
                        if column is None:
                            continue
                        carriers = unique_carriers(
                            [worksheet.cell(row_index, column).value]
                        )
                        if not carriers:
                            continue
                        by_month[(month, container, group)].extend(carriers)
                        all_months[(container, group)].extend(carriers)
        finally:
            workbook.close()

        resolved: dict[int, str] = {}
        for index, container, group, invoice_month in pending:
            lookup_months = _lookup_months(invoice_month)
            if lookup_months is None:
                candidates = unique_carriers(all_months.get((container, group), ()))
            else:
                candidates = unique_carriers(
                    carrier
                    for month in lookup_months
                    for carrier in by_month.get((month, container, group), ())
                )
            if len(candidates) == 1:
                resolved[index] = candidates[0]
        return resolved


__all__ = ["ReviewCarrierLookup"]
