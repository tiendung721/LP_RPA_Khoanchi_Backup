"""Đọc khối tổng hợp BK và tạo request JSON ổn định cho PAD."""

from __future__ import annotations

import json
import math
import os
import sys
import tempfile
from collections import OrderedDict
from collections.abc import Callable, Iterable
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

from app.services.excel.headers import normalize_header
from app.services.excel.payment_sync import (
    ArithmeticFormulaEvaluator,
    PaymentSyncError,
    find_summary_start,
)
from app.services.excel.posting import (
    FEE_HEADER_ALIASES,
    FEE_INVOICE_HEADER_ALIASES,
    INVOICE_NUMBER_HEADER_NAMES,
)
from app.services.excel.resolvers import MonthSheetService
from app.services.excel.workbook import (
    ExcelLockService,
    WorkbookGateway,
    ensure_supported_workbook,
)

from .contracts import (
    RPA_EXPENSE_OPERATION,
    RPA_STATUS_IMPORTED,
    RPA_STATUS_NOT_IMPORTED,
    RPA_FEE_GROUPS,
    PreparedRpaSelection,
    RpaExpenseAmounts,
    RpaExpensePlan,
    RpaFeeEntry,
    RpaSheetCandidate,
    RpaSqtItem,
)


ProgressCallback = Callable[[str], None] | None

SUMMARY_HEADERS: tuple[str, ...] = (
    "QT",
    "CƯỚC MB",
    "N.HẠ MB",
    "CƯỚC BIỂN",
    "N.HA VS D/O LỆNH",
    "CƯỚC MN",
    "Lưu cont",
    "Sửa chữa Cont",
    "QUÁ TẢI",
    "LƯU CONT/QUÁ TẢI",
)
SUMMARY_KEYS: tuple[str, ...] = (
    "sqt",
    "cuoc_bo_dong_hang",
    "nang_ha_dong_hang",
    "cuoc_bien",
    "nang_do_vs_lam_lenh",
    "cuoc_bo_tra_hang",
    "luu_cont",
    "sua_chua_cont",
    "qua_tai",
    "luu_cont_qua_tai",
)
STATUS_HEADER = "Trạng thái RPA"


class RpaExpenseError(RuntimeError):
    """Lỗi nghiệp vụ an toàn để hiển thị trực tiếp."""


def _progress(callback: ProgressCallback, message: str) -> None:
    if callback is not None:
        callback(message)


def normalize_sqt(value: Any) -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, int):
        return str(value) if value > 0 else ""
    if isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer() or value <= 0:
            return ""
        return str(int(value))
    text = str(value).strip()
    if not text:
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text if text.isdigit() and int(text) > 0 else ""


def normalize_rpa_status(value: Any) -> str:
    return (
        RPA_STATUS_IMPORTED
        if normalize_header(value) == normalize_header(RPA_STATUS_IMPORTED)
        else RPA_STATUS_NOT_IMPORTED
    )


class RpaExpenseService:
    def __init__(
        self,
        settings: Any,
        *,
        gateway: WorkbookGateway | None = None,
        lock_service: ExcelLockService | None = None,
        month_service: MonthSheetService | None = None,
        tracking_repository: Any | None = None,
    ) -> None:
        self.gateway = gateway or WorkbookGateway()
        self.lock_service = lock_service or ExcelLockService()
        self.months = month_service or MonthSheetService()
        self.tracking_repository = tracking_repository
        self.update_settings(settings)

    def update_settings(self, settings: Any) -> None:
        self.settings = settings
        self.bk_path = Path(str(getattr(settings, "bk_workbook_path", "") or ""))
        paths = getattr(settings, "paths", None)
        system_dir = Path(
            getattr(paths, "system_dir", Path("Output") / "_system")
        )
        self.runtime_dir = Path(
            getattr(paths, "rpa_dir", system_dir / "RPA")
        )

    def sheet_candidates(
        self, progress_callback: ProgressCallback = None
    ) -> list[RpaSheetCandidate]:
        target = self._target_path()
        _progress(progress_callback, "Đang đọc danh sách sheet BK…")
        self.lock_service.ensure_readable(target)
        workbook = self.gateway.load(target, read_only=True, data_only=False)
        try:
            candidates = []
            for name in workbook.sheetnames:
                parsed = self.months.parse_target_sheet(name)
                if parsed is None:
                    continue
                month, year = parsed
                candidates.append(RpaSheetCandidate(name, month, year))
        finally:
            workbook.close()
        if not candidates:
            raise RpaExpenseError("File BK không có sheet tháng dạng TMM YY.")
        return sorted(
            candidates,
            key=lambda item: (item.year, item.month, item.sheet_name),
            reverse=True,
        )

    def analyze_sheet(
        self,
        sheet_name: str,
        progress_callback: ProgressCallback = None,
    ) -> RpaExpensePlan:
        target = self._target_path()
        if not str(sheet_name).strip():
            raise RpaExpenseError("Chưa chọn sheet BK.")
        _progress(progress_callback, f"Đang đọc dữ liệu {sheet_name}…")
        self.lock_service.ensure_readable(target)
        fingerprint = self.gateway.fingerprint(target)
        workbook = self.gateway.load(target, read_only=False, data_only=False)
        try:
            if sheet_name not in workbook.sheetnames:
                raise RpaExpenseError(f"Không tìm thấy sheet {sheet_name}.")
            worksheet = workbook[sheet_name]
            items = self._extract_items(worksheet)
        finally:
            workbook.close()
        self.gateway.assert_unchanged(
            target, fingerprint, label="File BK"
        )
        if not items:
            raise RpaExpenseError(
                f"Sheet {sheet_name} không có SQT hợp lệ trong cột QT."
            )
        _progress(
            progress_callback,
            f"Đã đọc {len(items)} SQT; {sum(item.can_run for item in items)} SQT có thể chạy.",
        )
        latest_bk_revisions: tuple[tuple[str, int], ...] = ()
        latest_pad_sqt: tuple[str, ...] = ()
        if self.tracking_repository is not None:
            snapshot = self.tracking_repository.snapshot(target, sheet_name)
            latest_bk_revisions = tuple(snapshot.latest_bk_revisions.items())
            latest_pad_sqt = tuple(snapshot.latest_pad_sqt)
        return RpaExpensePlan(
            target,
            sheet_name,
            fingerprint,
            items,
            latest_bk_revisions=latest_bk_revisions,
            latest_pad_sqt=latest_pad_sqt,
        )

    def analyze_all_sheets(
        self, progress_callback: ProgressCallback = None
    ) -> RpaExpensePlan:
        target = self._target_path()
        _progress(progress_callback, "Đang đọc tất cả sheet tháng trong BK…")
        self.lock_service.ensure_readable(target)
        fingerprint = self.gateway.fingerprint(target)
        workbook = self.gateway.load(target, read_only=False, data_only=False)
        items: list[RpaSqtItem] = []
        sheet_names: list[str] = []
        try:
            for name in workbook.sheetnames:
                if self.months.parse_target_sheet(name) is None:
                    continue
                sheet_names.append(name)
                _progress(progress_callback, f"Đang đọc dữ liệu {name}…")
                try:
                    items.extend(self._extract_items(workbook[name]))
                except RpaExpenseError as exc:
                    raise RpaExpenseError(f"Sheet {name}: {exc}") from exc
        finally:
            workbook.close()
        self.gateway.assert_unchanged(target, fingerprint, label="File BK")
        if not sheet_names:
            raise RpaExpenseError("File BK không có sheet tháng dạng TMM YY.")
        if not items:
            raise RpaExpenseError("Các sheet tháng trong BK không có SQT hợp lệ.")

        # PAD tìm trên web bằng SQT, nên một SQT ở hai sheet không thể chọn an toàn.
        sheets_by_sqt: dict[str, set[str]] = {}
        for item in items:
            sheets_by_sqt.setdefault(item.sqt, set()).add(item.sheet_name)
        items = [
            replace(
                item,
                errors=(
                    *item.errors,
                    "SQT trùng ở nhiều sheet BK; PAD chỉ tìm bằng SQT.",
                ),
            )
            if len(sheets_by_sqt[item.sqt]) > 1
            else item
            for item in items
        ]
        latest_bk_revisions: list[tuple[str, int]] = []
        latest_pad_sqt: list[str] = []
        if self.tracking_repository is not None:
            for name in sheet_names:
                snapshot = self.tracking_repository.snapshot(target, name)
                latest_bk_revisions.extend(snapshot.latest_bk_revisions.items())
                latest_pad_sqt.extend(snapshot.latest_pad_sqt)
        _progress(
            progress_callback,
            f"Đã đọc {len(items)} SQT trên {len(sheet_names)} sheet; "
            f"{sum(item.can_run for item in items)} SQT có thể chạy.",
        )
        return RpaExpensePlan(
            target,
            "",
            fingerprint,
            tuple(items),
            latest_bk_revisions=tuple(latest_bk_revisions),
            latest_pad_sqt=tuple(latest_pad_sqt),
        )

    def _extract_items(self, worksheet: Any) -> tuple[RpaSqtItem, ...]:
        columns = self._summary_columns(worksheet)
        invoice_columns = self._invoice_columns(worksheet)
        fee_columns = self._fee_invoice_columns(worksheet)
        paired_invoice_columns = {
            invoice_column
            for _fee_column, invoice_column in fee_columns.values()
            if invoice_column is not None
        }
        status_column = self._optional_status_column(worksheet)
        evaluator = ArithmeticFormulaEvaluator(worksheet)
        groups: "OrderedDict[str, dict[str, Any]]" = OrderedDict()
        for row_number in range(2, int(worksheet.max_row or 0) + 1):
            try:
                sqt_value = evaluator.value(
                    worksheet.cell(row_number, columns["sqt"])
                )
            except PaymentSyncError:
                continue
            sqt = normalize_sqt(sqt_value)
            if not sqt:
                continue
            group = groups.setdefault(
                sqt,
                {
                    "rows": [],
                    "statuses": [],
                    "invoice_numbers": [],
                    "fee_entries": [],
                    "amounts": {
                        key: 0
                        for key in SUMMARY_KEYS
                        if key not in {"sqt", "luu_cont", "qua_tai"}
                    },
                    "errors": [],
                },
            )
            group["rows"].append(row_number)
            invoice_values: dict[int, str] = {}
            for column in invoice_columns:
                cell = worksheet.cell(row_number, column)
                invoice_value = cell.value
                if cell.data_type == "f":
                    try:
                        invoice_value = evaluator.value(cell)
                    except PaymentSyncError:
                        continue
                invoice = self._invoice_text(invoice_value)
                invoice_values[column] = invoice
                if invoice and not any(
                    value.casefold() == invoice.casefold()
                    for value in group["invoice_numbers"]
                ):
                    group["invoice_numbers"].append(invoice)
            for category_key, _category_label, fee_specs in RPA_FEE_GROUPS:
                for fee_key, fee_label in fee_specs:
                    pair = fee_columns.get(fee_key)
                    if pair is None:
                        continue
                    fee_column, invoice_column = pair
                    try:
                        fee_amount = self._money_value(
                            evaluator, worksheet.cell(row_number, fee_column)
                        )
                    except RpaExpenseError:
                        fee_amount = None
                    invoice = invoice_values.get(invoice_column, "")
                    if fee_amount in (0, None) and not invoice:
                        continue
                    group["fee_entries"].append(
                        RpaFeeEntry(
                            category_key=category_key,
                            fee_key=fee_key,
                            fee_label=fee_label,
                            source_row=row_number,
                            amount=fee_amount,
                            invoice_number=invoice,
                            has_invoice_column=invoice_column is not None,
                        )
                    )
            for invoice_column, invoice in invoice_values.items():
                if invoice and invoice_column not in paired_invoice_columns:
                    header = normalize_header(
                        worksheet.cell(1, invoice_column).value
                    )
                    group["fee_entries"].append(
                        RpaFeeEntry(
                            category_key="other",
                            fee_key="OTHER",
                            fee_label=(
                                "Phí Seal (không gửi PAD)"
                                if header == normalize_header("Số HĐ Seal")
                                else "HĐ chưa gắn phí"
                            ),
                            source_row=row_number,
                            amount=None,
                            invoice_number=invoice,
                            has_invoice_column=True,
                        )
                    )
            status_value = (
                worksheet.cell(row_number, status_column).value
                if status_column is not None
                else None
            )
            group["statuses"].append(normalize_rpa_status(status_value))
            for key in (
                "cuoc_bo_dong_hang",
                "nang_ha_dong_hang",
                "cuoc_bien",
                "nang_do_vs_lam_lenh",
                "cuoc_bo_tra_hang",
                "luu_cont_qua_tai",
                "sua_chua_cont",
            ):
                cell = worksheet.cell(row_number, columns[key])
                try:
                    amount = self._money_value(evaluator, cell)
                except RpaExpenseError as exc:
                    group["errors"].append(f"{cell.coordinate}: {exc}")
                    continue
                group["amounts"][key] += amount
        return tuple(
            self._build_item(sqt, value, worksheet.title)
            for sqt, value in groups.items()
        )

    def prepare_selection(
        self,
        plan: RpaExpensePlan,
        selected_sqt: Iterable[str],
        progress_callback: ProgressCallback = None,
    ) -> PreparedRpaSelection:
        selected = list(dict.fromkeys(str(value).strip() for value in selected_sqt))
        selected = [value for value in selected if value]
        if not selected:
            raise RpaExpenseError("Vui lòng chọn ít nhất một SQT.")
        lookup = plan.item_map()
        unknown = [value for value in selected if value not in lookup]
        if unknown:
            raise RpaExpenseError(
                "Danh sách SQT đã thay đổi; không còn tìm thấy: "
                + ", ".join(unknown)
            )
        blocked = [lookup[value] for value in selected if not lookup[value].can_run]
        if blocked:
            detail = "; ".join(
                f"{item.sqt}: {item.validation_message}" for item in blocked
            )
            raise RpaExpenseError(
                "Có SQT chưa đủ điều kiện chạy RPA: " + detail
            )
        selected_sheets = {
            lookup[value].sheet_name or plan.sheet_name for value in selected
        }
        request_sheet = next(iter(selected_sheets)) if len(selected_sheets) == 1 else None
        self.gateway.assert_unchanged(
            plan.bk_path, plan.fingerprint, label="File BK"
        )
        # Chỉ preflight. Không giữ khóa vì PAD cần cập nhật trạng thái sau đó.
        with self.lock_service.acquire(plan.bk_path):
            pass
        _progress(progress_callback, "Đang tạo dữ liệu đầu vào cho PAD…")
        run_id = (
            datetime.now().strftime("%Y%m%d_%H%M%S")
            + "_"
            + uuid4().hex[:8]
        )
        self.runtime_dir.mkdir(parents=True, exist_ok=True)
        # Flow PAD cũ đọc trực tiếp một file cố định. Giữ đúng cơ chế đó để
        # người dùng không phải khai báo Input variable hoặc truyền tham số.
        selection_path = self.runtime_dir / "rpa_input_selection.json"
        payload: dict[str, Any] = {
            "version": 1,
            "operation": RPA_EXPENSE_OPERATION,
            "run_id": run_id,
            "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "project_root": str(
                Path(getattr(self.settings, "data_root", Path.cwd())).resolve()
            ),
            "bk_file": str(plan.bk_path.resolve()),
            "sheet_name": request_sheet,
            "source_fingerprint": plan.fingerprint.to_dict(),
            "items": [
                {
                    **lookup[value].to_payload(),
                    **(
                        {"bk_revision": revision}
                        if self.tracking_repository is not None
                        and (revision := self.tracking_repository.bk_revision(
                            plan.bk_path, lookup[value].sheet_name or plan.sheet_name, value
                        )) is not None
                        else {}
                    ),
                }
                for value in selected
            ],
            "status_callback": {
                "when": "AFTER_WEB_SAVE_SUCCESS",
                "status": RPA_STATUS_IMPORTED,
                "python_executable": str(Path(sys.executable).resolve()),
                "script": str(
                    (
                        Path(__file__).resolve().parents[2]
                        / "scripts"
                        / "rpa_excel_helper.py"
                    ).resolve()
                ),
                "arguments": [
                    "mark-imported",
                    "--selection",
                    str(selection_path.resolve()),
                    "--sqt",
                    "{sqt}",
                ],
            },
        }
        self._write_json_atomic(selection_path, payload)
        return PreparedRpaSelection(
            selection_path=selection_path.resolve(),
            run_id=run_id,
            item_count=len(selected),
            payload=payload,
        )

    @property
    def latest_launched_path(self) -> Path:
        return self.runtime_dir / "rpa_latest_launched.json"

    def record_launched(self, prepared: PreparedRpaSelection) -> Path:
        """Chỉ thay snapshot xem lại sau khi BAT/PAD đã khởi chạy thành công."""

        payload = dict(prepared.payload)
        launched_at = datetime.now().astimezone().isoformat(timespec="seconds")
        payload["launched_at"] = launched_at
        self.runtime_dir.mkdir(parents=True, exist_ok=True)
        self._write_json_atomic(self.latest_launched_path, payload)
        if self.tracking_repository is not None:
            selected_by_sheet: dict[str, list[str]] = {}
            for item in payload["items"]:
                sheet = str(item.get("sheet_name") or payload["sheet_name"] or "")
                selected_by_sheet.setdefault(sheet, []).append(str(item["sqt"]))
            for sheet, sqt_values in selected_by_sheet.items():
                self.tracking_repository.record_latest_pad(
                    payload["bk_file"],
                    sheet,
                    sqt_values,
                    run_id=str(payload["run_id"]),
                    launched_at=launched_at,
                )
        return self.latest_launched_path.resolve()

    def load_latest_launched(self) -> dict[str, Any] | None:
        path = self.latest_launched_path
        if not path.is_file():
            return None
        try:
            payload = json.loads(path.read_text(encoding="utf-8-sig"))
        except (OSError, UnicodeError, json.JSONDecodeError) as exc:
            raise RpaExpenseError(
                f"Không đọc được dữ liệu RPA gần nhất: {exc}"
            ) from exc
        if not isinstance(payload, dict):
            raise RpaExpenseError("Dữ liệu RPA gần nhất phải là object.")
        if payload.get("operation") != RPA_EXPENSE_OPERATION:
            raise RpaExpenseError("Dữ liệu gần nhất không đúng nghiệp vụ RPA.")
        if not isinstance(payload.get("items"), list):
            raise RpaExpenseError("Dữ liệu RPA gần nhất thiếu danh sách SQT.")
        return payload

    def _target_path(self) -> Path:
        if not str(self.bk_path).strip() or str(self.bk_path) == ".":
            raise RpaExpenseError("Chưa cấu hình file BK Tổng hợp.")
        target = ensure_supported_workbook(self.bk_path)
        if not target.is_file():
            raise RpaExpenseError(f"Không tìm thấy file BK: {target}")
        return target

    @staticmethod
    def _summary_columns(worksheet: Any) -> dict[str, int]:
        start = find_summary_start(worksheet)
        if start is None:
            raise RpaExpenseError(
                "Không nhận diện được khối cột tổng hợp bắt đầu bằng QT."
            )
        columns = {
            key: start + offset
            for offset, key in enumerate(SUMMARY_KEYS)
        }
        actual = tuple(
            normalize_header(worksheet.cell(1, start + offset).value)
            for offset in range(len(SUMMARY_HEADERS))
        )
        expected = tuple(normalize_header(value) for value in SUMMARY_HEADERS)
        if actual != expected:
            differences = [
                f"{SUMMARY_HEADERS[index]} → "
                f"{worksheet.cell(1, start + index).value!r}"
                for index in range(len(expected))
                if actual[index] != expected[index]
            ]
            raise RpaExpenseError(
                "Khối cột tổng hợp BK không đúng cấu trúc: "
                + "; ".join(differences)
            )
        return columns

    @staticmethod
    def _invoice_columns(worksheet: Any) -> tuple[int, ...]:
        summary_start = find_summary_start(worksheet)
        if summary_start is None:
            return ()
        return tuple(
            column
            for column in range(1, summary_start)
            if normalize_header(worksheet.cell(1, column).value)
            in INVOICE_NUMBER_HEADER_NAMES
        )

    @staticmethod
    def _fee_invoice_columns(
        worksheet: Any,
    ) -> dict[str, tuple[int, int | None]]:
        summary_start = find_summary_start(worksheet)
        if summary_start is None:
            return {}
        headers = {
            column: normalize_header(worksheet.cell(1, column).value)
            for column in range(1, summary_start)
        }
        notes_columns = [
            column for column, value in headers.items()
            if value == normalize_header("GHI CHÚ")
        ]
        notes_column = min(notes_columns) if notes_columns else None
        result: dict[str, tuple[int, int | None]] = {}
        for _category_key, _category_label, fee_specs in RPA_FEE_GROUPS:
            for fee_key, _fee_label in fee_specs:
                aliases = {
                    normalize_header(value)
                    for value in FEE_HEADER_ALIASES[fee_key]
                }
                matches = [
                    column for column, value in headers.items()
                    if value in aliases
                    and not (
                        fee_key == "CBDH"
                        and value == normalize_header("ĐƠN GIÁ")
                        and notes_column is not None
                        and column >= notes_column
                    )
                ]
                if len(matches) != 1:
                    continue
                fee_column = matches[0]
                invoice_column: int | None = None
                if fee_key in FEE_INVOICE_HEADER_ALIASES:
                    invoice_aliases = {
                        normalize_header(value)
                        for value in FEE_INVOICE_HEADER_ALIASES[fee_key]
                    }
                    invoice_matches = [
                        column for column, value in headers.items()
                        if value in invoice_aliases
                    ]
                    if len(invoice_matches) == 1:
                        invoice_column = invoice_matches[0]
                elif fee_key != "LL" and (
                    headers.get(fee_column + 1) in INVOICE_NUMBER_HEADER_NAMES
                ):
                    invoice_column = fee_column + 1
                result[fee_key] = fee_column, invoice_column
        return result

    @staticmethod
    def _invoice_text(value: Any) -> str:
        if value is None or isinstance(value, bool):
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    @staticmethod
    def _optional_status_column(worksheet: Any) -> int | None:
        expected = normalize_header(STATUS_HEADER)
        matches = [
            column
            for column in range(1, int(worksheet.max_column or 0) + 1)
            if normalize_header(worksheet.cell(1, column).value) == expected
        ]
        if len(matches) > 1:
            raise RpaExpenseError(
                "Có nhiều cột Trạng thái RPA; không thể chọn an toàn."
            )
        return matches[0] if matches else None

    @staticmethod
    def _money_value(
        evaluator: ArithmeticFormulaEvaluator, cell: Any
    ) -> int:
        try:
            value = evaluator.value(cell)
        except PaymentSyncError as exc:
            raise RpaExpenseError(str(exc)) from exc
        if value in (None, ""):
            return 0
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise RpaExpenseError("Giá trị không phải số tiền.")
        if isinstance(value, float) and not value.is_integer():
            raise RpaExpenseError("Số tiền phải là số nguyên.")
        amount = int(value)
        if amount < 0:
            raise RpaExpenseError("Số tiền không được âm.")
        return amount

    @staticmethod
    def _build_item(
        sqt: str, group: dict[str, Any], sheet_name: str
    ) -> RpaSqtItem:
        statuses = list(group["statuses"])
        status = (
            RPA_STATUS_IMPORTED
            if statuses and all(value == RPA_STATUS_IMPORTED for value in statuses)
            else RPA_STATUS_NOT_IMPORTED
        )
        amounts = RpaExpenseAmounts(**group["amounts"])
        return RpaSqtItem(
            sqt=sqt,
            source_rows=tuple(group["rows"]),
            status=status,
            amounts=amounts,
            invoice_numbers=tuple(group["invoice_numbers"]),
            errors=tuple(dict.fromkeys(group["errors"])),
            sheet_name=sheet_name,
            fee_entries=tuple(group["fee_entries"]),
        )

    @staticmethod
    def _write_json_atomic(path: Path, payload: dict[str, Any]) -> None:
        data = (
            json.dumps(
                payload,
                ensure_ascii=False,
                indent=2,
                allow_nan=False,
            )
            + "\n"
        ).encode("utf-8")
        temporary: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="wb",
                prefix=f".{path.name}.",
                suffix=".tmp",
                dir=path.parent,
                delete=False,
            ) as handle:
                temporary = Path(handle.name)
                handle.write(data)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary, path)
            temporary = None
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)


__all__ = [
    "RpaExpenseError",
    "RpaExpenseService",
    "STATUS_HEADER",
    "SUMMARY_HEADERS",
    "normalize_rpa_status",
    "normalize_sqt",
]
