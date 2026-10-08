"""Export a persisted, one-sheet audit of a completed expense posting run."""

from __future__ import annotations

import os
import tempfile
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .models import ExcelOperation, PostingPlan, PostingResult
from .workbook import WorkbookChangedError, workbook_fingerprint


class PostingExportError(RuntimeError):
    pass


HEADERS = (
    "STT", "File nguồn", "Mã chứng từ nguồn", "Dòng nguồn", "Số HĐ nguồn",
    "Ngày HĐ", "Số cont trên HĐ", "Căn cứ số cont", "Container", "B/L",
    "SQT", "Tàu/chuyến", "Tên tàu", "Số chuyến", "Mã phí nguồn",
    "Mã phí ghi BK", "Quy tắc tính", "Bên VT nguồn", "Bên VT hiệu lực",
    "Số tiền nguồn", "Giá trị ô phí BK sau xử lý", "Tiền ghi lượt này",
    "Số HĐ trong BK sau xử lý", "Sheet BK", "Dòng BK", "Ô phí BK",
    "Kết quả tiền", "Kết quả Số HĐ", "Kết quả bên VT", "Trạng thái chung",
    "Ghi chú",
)
ROW_FIELDS = (
    "ordinal", "source_document_name", "source_document_id", "source_row",
    "invoice_no", "invoice_date", "invoice_container_count", "container_count_basis",
    "container", "bl", "sqt", "vessel_voyage_raw", "vessel_name", "voyage_no",
    "fee", "fee_selected", "rule", "carrier", "carrier_effective", "amount",
    "value_after", "amount_written_this_run", "invoice_value_after", "sheet_name",
    "target_row", "target_cell", "amount_status", "invoice_status", "carrier_status",
    "status", "note",
)
assert len(HEADERS) == len(ROW_FIELDS)
HEADER_ROW = 7
FIRST_DATA_ROW = HEADER_ROW + 1
AMOUNT_COLUMNS = (20, 21, 22)
NUMBER_COLUMNS = (1, 4, 7, 11, 25)


def _code(value: Any) -> str:
    return str(getattr(value, "value", value) or "").split(".")[-1].upper()


def _invoice_key(row: Mapping[str, Any]) -> tuple[str, str, str]:
    document = str(row.get("source_document_id") or "")
    invoice = str(row.get("invoice_no") or "").strip().casefold()
    return (document, invoice or "__NO_INVOICE__", str(row.get("invoice_date") or ""))


def _component_status(action: Mapping[str, Any], component: str) -> str:
    if action.get(f"{component}_write"):
        return "Đã ghi"
    choice = _code(action.get("action" if component == "amount" else f"{component}_action"))
    if choice in {"KEEP_EXISTING", "KEEP_FORMULA", "SKIP_INVOICE"}:
        return "Giữ nguyên"
    if choice == "SKIP":
        return "Bỏ qua"
    if component == "invoice" and not action.get("invoice_selected"):
        return "Không có"
    if component == "carrier" and not action.get("carrier_group"):
        return "Không áp dụng"
    return "Không đổi"


def _written_amount(action: Mapping[str, Any]) -> int | float:
    after = action.get("value_after")
    if _code(action.get("action")) == "ADD":
        try:
            value = Decimal(str(after)) - Decimal(str(action.get("value_before")))
        except Exception as exc:
            raise PostingExportError("Không tính được số tiền điều chỉnh vừa ghi vào BK.") from exc
    else:
        try:
            value = Decimal(str(after))
        except Exception as exc:
            raise PostingExportError("Số tiền ghi vào BK không hợp lệ.") from exc
    if not value.is_finite():
        raise PostingExportError("Số tiền ghi vào BK không hữu hạn.")
    return int(value) if value == value.to_integral_value() else float(value)


def build_posting_export_snapshot(
    plan: PostingPlan,
    actions: Sequence[Mapping[str, Any]],
    result: PostingResult,
) -> dict[str, Any]:
    """Retain every confirmed source item, including skipped and old items."""

    by_index: dict[int, Mapping[str, Any]] = {}
    first_in_group: set[int] = set()
    for action in actions:
        indices = [int(value) for value in action.get("source_indices", ())]
        if indices:
            first_in_group.add(indices[0])
        for index in indices:
            if index in by_index:
                raise PostingExportError(f"Dòng nguồn {index + 1} được xử lý hai lần.")
            by_index[index] = action
    previous = {
        int(item["source_item_index"]): item
        for item in plan.previously_posted_items
        if item.get("source_item_index") is not None
    }
    rows: list[dict[str, Any]] = []
    for ordinal, raw in enumerate(plan.source_rows, 1):
        source = dict(raw)
        index = int(source.get("source_item_index", ordinal - 1))
        action = by_index.get(index)
        old = previous.get(index, {})
        row = {
            "ordinal": ordinal,
            "source_row": index + 1,
            "source_document_name": source.get("source_document_name") or "Không có tên file nguồn",
            "source_document_id": source.get("source_document_id"),
            "invoice_no": source.get("invoice_no"),
            "invoice_date": source.get("invoice_date"),
            "invoice_container_count": source.get("invoice_container_count"),
            "container_count_basis": source.get("container_count_basis"),
            "container": source.get("container"),
            "bl": source.get("bl"),
            "sqt": source.get("input_sqt"),
            "vessel_voyage_raw": source.get("vessel_voyage_raw"),
            "vessel_name": source.get("vessel_name"),
            "voyage_no": source.get("voyage_no"),
            "fee": source.get("fee"),
            "fee_selected": source.get("fee"),
            "rule": source.get("rule"),
            "carrier": source.get("carrier"),
            "carrier_effective": None,
            "amount": source.get("amount"),
            "value_after": None,
            "amount_written_this_run": None,
            "invoice_value_after": None,
            "sheet_name": old.get("sheet_name"),
            "target_row": old.get("target_row"),
            "target_cell": old.get("target_cell"),
            "amount_status": "Đã nhập trước" if index in plan.already_posted_indices else "Chưa xử lý",
            "invoice_status": "—",
            "carrier_status": "—",
            "status": "Đã nhập trước" if index in plan.already_posted_indices else "Chưa xử lý",
            "note": "Đã được ghi trong lượt trước." if index in plan.already_posted_indices else "",
        }
        if action is not None:
            status = _code(action.get("status"))
            amount_write = bool(action.get("amount_write"))
            invoice_write = bool(action.get("invoice_write"))
            carrier_write = bool(action.get("carrier_write"))
            grouped = len(action.get("source_indices", ())) > 1
            if amount_write:
                label = "Đã ghi tiền"
            elif invoice_write or carrier_write:
                label = "Chỉ cập nhật thông tin"
            elif status == "ALREADY_EXISTS":
                label = "Đã có trong BK"
            elif status == "NOT_MATCHED":
                label = "Không tìm thấy dòng BK"
            elif status == "UNRESOLVED":
                label = "Chưa xác định vị trí BK"
            else:
                choice = _code(action.get("action"))
                label = (
                    "Giữ nguyên" if choice in {"KEEP_EXISTING", "KEEP_FORMULA"}
                    else "Bỏ qua" if status == "USER_SKIPPED" or choice == "SKIP"
                    else "Giữ nguyên"
                )
            notes: list[str] = []
            if grouped:
                notes.append("Gộp cùng ô BK; tiền ghi lượt này chỉ tính một lần.")
            if not amount_write and label == "Bỏ qua":
                notes.append("Khoản phí không được ghi vào BK.")
            row.update(
                fee_selected=action.get("fee_selected") or row["fee"],
                carrier_effective=action.get("carrier_effective") or action.get("carrier_value_after"),
                sqt=action.get("source_sqt") or row["sqt"],
                value_after=action.get("value_after"),
                amount_written_this_run=(
                    _written_amount(action) if amount_write and index in first_in_group else None
                ),
                invoice_value_after=action.get("invoice_value_after"),
                sheet_name=action.get("sheet_name"),
                target_row=action.get("target_row"),
                target_cell=action.get("target_cell"),
                amount_status=(
                    "Gộp cùng ô BK" if amount_write and grouped and index not in first_in_group
                    else "Đã ghi" if amount_write else _component_status(action, "amount")
                ),
                invoice_status=_component_status(action, "invoice"),
                carrier_status=_component_status(action, "carrier"),
                status=label,
                note=" ".join(notes),
            )
        rows.append(row)
    missing = set(by_index).difference(
        int(source.get("source_item_index", position))
        for position, source in enumerate(plan.source_rows)
    )
    if missing:
        raise PostingExportError("Kết quả ghi BK có dòng không thuộc nguồn đã xác nhận.")
    rows.sort(key=lambda row: (
        str(row.get("source_document_name") or "").casefold(),
        str(row.get("source_document_id") or "").casefold(),
        str(row.get("invoice_no") or "").casefold(),
        int(row["source_row"]),
    ))
    for ordinal, row in enumerate(rows, 1):
        row["ordinal"] = ordinal
    total = sum(
        (Decimal(str(row["amount_written_this_run"])) for row in rows
         if row["amount_written_this_run"] is not None),
        Decimal(0),
    )
    return {
        "version": 1,
        "run_id": result.run_id,
        "completed_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "source_path": str(plan.batch_path),
        "bk_path": str(result.target_path),
        "bk_sheets": list(result.target_sheets or ([result.sheet_name] if result.sheet_name else [])),
        "invoice_count": len({_invoice_key(row) for row in rows}),
        "row_count": len(rows),
        "written_total": str(total),
        "rows": rows,
    }


@dataclass(frozen=True, slots=True)
class PostingExportPreview:
    run_id: int
    completed_at: str
    row_count: int
    invoice_count: int
    written_total: Decimal
    filename: str
    last_export_path: Path | None = None


@dataclass(frozen=True, slots=True)
class PostingExportResult:
    target_path: Path
    run_id: int
    row_count: int
    invoice_count: int
    written_total: Decimal
    operation: ExcelOperation = ExcelOperation.POSTING_EXPORT

    @property
    def message(self) -> str:
        return f"{self.invoice_count} HĐ, {self.row_count} dòng → {self.target_path}"


class PostingExportService:
    def __init__(self, run_repository: Any, posting_repository: Any) -> None:
        self.run_repository = run_repository
        self.posting_repository = posting_repository

    def latest_preview(self) -> PostingExportPreview | None:
        run = self.run_repository.get_latest(
            operation=ExcelOperation.EXPENSE_POSTING,
            statuses=("SUCCEEDED", "NO_CHANGES"),
        )
        if run is None:
            return None
        snapshot = self.posting_repository.get_export_snapshot(run.id)
        if snapshot is None:
            return None
        completed = str(snapshot["completed_at"])
        stamp = datetime.fromisoformat(completed.replace("Z", "+00:00")).strftime("%Y%m%d_%H%M%S")
        return PostingExportPreview(
            run_id=run.id,
            completed_at=completed,
            row_count=int(snapshot["row_count"]),
            invoice_count=int(snapshot["invoice_count"]),
            written_total=Decimal(str(snapshot["written_total"])),
            filename=f"Ket_qua_nhap_BK_{stamp}_L{run.id}.xlsx",
            last_export_path=self.posting_repository.get_export_path(run.id),
        )

    def export(
        self,
        run_id: int,
        destination: str | Path,
        *,
        overwrite: bool = False,
        progress_callback: Callable[[str], None] | None = None,
    ) -> PostingExportResult:
        run = self.run_repository.require_by_id(run_id)
        if run.operation != ExcelOperation.EXPENSE_POSTING.value or run.status not in {"SUCCEEDED", "NO_CHANGES"}:
            raise PostingExportError("Lượt nhập BK chưa hoàn tất.")
        snapshot = self.posting_repository.get_export_snapshot(run_id)
        if snapshot is None:
            raise PostingExportError("Lượt nhập này chưa có dữ liệu đầy đủ để xuất.")
        target = Path(destination)
        if target.suffix.lower() != ".xlsx":
            raise PostingExportError("File xuất phải có đuôi .xlsx.")
        if run.target_path is not None and target.resolve() == Path(run.target_path).resolve():
            raise PostingExportError("Không thể ghi đè file BK.")
        before = workbook_fingerprint(target) if target.exists() else None
        if before is not None and not overwrite:
            raise FileExistsError(f"File đã tồn tại: {target}")
        target.parent.mkdir(parents=True, exist_ok=True)
        temp_path: Path | None = None
        try:
            if progress_callback:
                progress_callback("Đang dựng bảng chi phí vừa nhập…")
            book = self._workbook(snapshot)
            handle, name = tempfile.mkstemp(prefix=".posting_export_", suffix=".xlsx", dir=target.parent)
            os.close(handle)
            temp_path = Path(name)
            try:
                book.save(temp_path)
            finally:
                book.close()
            self._verify(temp_path, snapshot)
            if before is None:
                if target.exists():
                    raise FileExistsError(f"File đích vừa xuất hiện: {target}")
            elif workbook_fingerprint(target) != before:
                raise WorkbookChangedError(label="File đích", path=target)
            os.replace(temp_path, target)
            temp_path = None
            self.posting_repository.set_export_path(run_id, target)
            return PostingExportResult(
                target, run_id, int(snapshot["row_count"]),
                int(snapshot["invoice_count"]), Decimal(str(snapshot["written_total"])),
            )
        finally:
            if temp_path is not None:
                temp_path.unlink(missing_ok=True)

    @staticmethod
    def _workbook(snapshot: Mapping[str, Any]) -> Workbook:
        book = Workbook()
        sheet = book.active
        sheet.title = "Chi phí nhập BK"
        end = get_column_letter(len(HEADERS))
        counts = Counter(str(row.get("status") or "") for row in snapshot["rows"])
        count_text = "  |  ".join(
            f"{label}: {counts[label]}" for label in (
                "Đã ghi tiền", "Chỉ cập nhật thông tin", "Đã có trong BK",
                "Giữ nguyên", "Bỏ qua", "Đã nhập trước", "Không tìm thấy dòng BK",
                "Chưa xác định vị trí BK",
            ) if counts[label]
        )
        for row_number, value in (
            (1, "KẾT QUẢ NHẬP KHOẢN CHI VÀO BK"),
            (2, f"Lượt nhập #{snapshot['run_id']}  |  Hoàn tất: {snapshot['completed_at']}"),
            (3, f"File BK: {snapshot['bk_path']}  |  Sheet: {', '.join(snapshot['bk_sheets'])}"),
            (4, f"Nguồn: {snapshot['source_path']}"),
            (5, f"{snapshot['invoice_count']} HĐ  |  {snapshot['row_count']} dòng  |  "
                f"Tổng tiền ghi lượt này: {Decimal(str(snapshot['written_total'])):,.0f} đ"),
            (6, f"Trạng thái: {count_text}"),
        ):
            sheet.merge_cells(f"A{row_number}:{end}{row_number}")
            cell = sheet.cell(row_number, 1, value)
            cell.font = Font(bold=row_number in {1, 5}, size=15 if row_number == 1 else 10,
                             color="FFFFFF" if row_number == 1 else "203247")
            if row_number == 1:
                cell.fill = PatternFill("solid", fgColor="205781")
            sheet.row_dimensions[row_number].height = 27 if row_number == 1 else 21
        for column, title in enumerate(HEADERS, 1):
            cell = sheet.cell(HEADER_ROW, column, title)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="205781")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            width = 18
            if column in {2, 3, 12, 18, 19, 31}:
                width = 28
            elif column in {20, 21, 22, 23, 30}:
                width = 23
            sheet.column_dimensions[get_column_letter(column)].width = width
        sheet.row_dimensions[HEADER_ROW].height = 46
        for row_number, data in enumerate(snapshot["rows"], FIRST_DATA_ROW):
            for column, field in enumerate(ROW_FIELDS, 1):
                value = data.get(field)
                cell = sheet.cell(row_number, column)
                if field == "invoice_date" and isinstance(value, str):
                    try:
                        value = date.fromisoformat(value)
                    except ValueError:
                        pass
                if value not in (None, ""):
                    cell.value = value
                if isinstance(value, date):
                    cell.number_format = "dd/mm/yyyy"
                elif isinstance(value, str):
                    cell.data_type = "s"
                    cell.number_format = "@"
                elif column in AMOUNT_COLUMNS:
                    cell.number_format = '#,##0;[Red](#,##0)'
                elif column in NUMBER_COLUMNS:
                    cell.number_format = "0"
                if row_number % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F3F7FA")
            if data.get("status") in {"Bỏ qua", "Không tìm thấy dòng BK", "Chưa xác định vị trí BK"}:
                sheet.cell(row_number, 30).font = Font(color="A16207", bold=True)
        sheet.freeze_panes = f"F{FIRST_DATA_ROW}"
        if snapshot["rows"]:
            sheet.auto_filter.ref = f"A{HEADER_ROW}:{end}{sheet.max_row}"
        sheet.print_title_rows = f"1:{HEADER_ROW}"
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        return book

    @staticmethod
    def _verify(path: Path, snapshot: Mapping[str, Any]) -> None:
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            if book.sheetnames != ["Chi phí nhập BK"]:
                raise PostingExportError("File xuất không có đúng một sheet.")
            sheet = book.active
            if tuple(cell.value for cell in sheet[HEADER_ROW]) != HEADERS:
                raise PostingExportError("Header file xuất không khớp format.")
            if sheet.max_row != HEADER_ROW + int(snapshot["row_count"]):
                raise PostingExportError("Số dòng trong file xuất không khớp dữ liệu nguồn.")
            total = sum(
                (Decimal(str(sheet.cell(row, 22).value or 0))
                 for row in range(FIRST_DATA_ROW, sheet.max_row + 1)),
                Decimal(0),
            )
            if total != Decimal(str(snapshot["written_total"])):
                raise PostingExportError("Tổng tiền file xuất không khớp lượt nhập BK.")
        finally:
            book.close()
