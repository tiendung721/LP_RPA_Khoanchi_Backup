"""Đọc khối tổng hợp BK và tạo request JSON ổn định cho PAD."""

from __future__ import annotations

import json
import math
import os
import sys
import tempfile
from collections import OrderedDict
from collections.abc import Callable, Iterable
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
    PreparedRpaSelection,
    RpaExpenseAmounts,
    RpaExpensePlan,
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
    ) -> None:
        self.gateway = gateway or WorkbookGateway()
        self.lock_service = lock_service or ExcelLockService()
        self.months = month_service or MonthSheetService()
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
            columns = self._summary_columns(worksheet)
            status_column = self._optional_status_column(worksheet)
            evaluator = ArithmeticFormulaEvaluator(worksheet)
            groups: "OrderedDict[str, dict[str, Any]]" = OrderedDict()
            first_row = 2
            max_row = int(worksheet.max_row or 0)
            for row_number in range(first_row, max_row + 1):
                try:
                    sqt_value = evaluator.value(
                        worksheet.cell(row_number, columns["sqt"])
                    )
                except PaymentSyncError:
                    # Dòng không có QT hợp lệ không thuộc danh sách RPA.
                    continue
                sqt = normalize_sqt(sqt_value)
                if not sqt:
                    continue
                group = groups.setdefault(
                    sqt,
                    {
                        "rows": [],
                        "statuses": [],
                        "amounts": {
                            key: 0
                            for key in SUMMARY_KEYS
                            if key
                            not in {"sqt", "luu_cont", "qua_tai"}
                        },
                        "errors": [],
                    },
                )
                group["rows"].append(row_number)
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
                        group["errors"].append(
                            f"{cell.coordinate}: {exc}"
                        )
                        continue
                    group["amounts"][key] += amount
            items = tuple(self._build_item(sqt, value) for sqt, value in groups.items())
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
        return RpaExpensePlan(target, sheet_name, fingerprint, items)

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
            "sheet_name": plan.sheet_name,
            "source_fingerprint": plan.fingerprint.to_dict(),
            "items": [lookup[value].to_payload() for value in selected],
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
        payload["launched_at"] = datetime.now().astimezone().isoformat(
            timespec="seconds"
        )
        self.runtime_dir.mkdir(parents=True, exist_ok=True)
        self._write_json_atomic(self.latest_launched_path, payload)
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
    def _build_item(sqt: str, group: dict[str, Any]) -> RpaSqtItem:
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
            errors=tuple(dict.fromkeys(group["errors"])),
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
