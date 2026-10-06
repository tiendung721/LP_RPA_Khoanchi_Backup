"""Read NAM fees from BK and export one carrier's detail to a standalone workbook."""

from __future__ import annotations

import os
import re
import tempfile
from collections import Counter
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Any, Callable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .models import ExcelOperation, PaymentSyncItem, SourceSheetCandidate, WorkbookFingerprint
from .payment_sync import (
    NAM_FIELDS, NormalizationReport, _ensure_bk_structure, _has_money,
    _source_target_items,
)
from .resolvers import MonthSheetService
from .workbook import WorkbookChangedError, WorkbookGateway, ensure_supported_workbook, workbook_fingerprint


class CarrierExportError(RuntimeError):
    pass


_MONTH_SUFFIX = re.compile(r"(?:\s+BK\s*T?\s*|\s+T\s*)(?:0?[1-9]|1[0-2])$", re.IGNORECASE)
_BAD_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_MULTIPLE_CARRIERS = re.compile(r"\s*(?:/|;|\n|\s+\+\s+)\s*")
_FIELDS = tuple(NAM_FIELDS)
_AMOUNT_COLUMNS = (4, 6, 8, 10, 12, 14, 15, 17, 19)
_HEADERS = (
    "Tháng BK", "QT", "SỐ CONT", "Nâng vỏ", "Số HĐ Nâng vỏ",
    "Nâng hàng", "Số HĐ Nâng hàng", "Hạ vỏ", "Số HĐ Hạ vỏ",
    "VS + D/O", "Số HĐ VS + D/O", "Seal", "Số HĐ Seal", "Làm lệnh",
    "Lưu cont", "Số HĐ Lưu cont", "Sửa chữa", "Số HĐ Sửa chữa",
    "Quá tải", "Số HĐ Quá tải", "N.HA VS D/O LỆNH",
    "Lưu cont + Quá tải", "Sửa chữa Cont", "Tổng tiền",
)


def normalize_carrier(value: Any) -> str | None:
    text = " ".join(str(value or "").split()).strip()
    if not text:
        return None
    parts = _MULTIPLE_CARRIERS.split(text)
    names = []
    for part in parts:
        if not part or part == "0":
            return None
        name = _MONTH_SUFFIX.sub("", part).strip(" ._-")
        if not name:
            return None
        names.append(name)
    return names[0] if len({name.casefold() for name in names}) == 1 else None


def _number(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal(0)
    number = Decimal(str(value))
    if not number.is_finite():
        raise ValueError("Số tiền không hữu hạn.")
    return number


@dataclass(frozen=True, slots=True)
class CarrierExportRow:
    sheet_name: str
    month: int
    year: int
    carrier_key: str | None
    item: PaymentSyncItem
    error: str | None = None


@dataclass(frozen=True, slots=True)
class CarrierCandidate:
    key: str
    name: str
    counts: tuple[int, ...]

    @property
    def total(self) -> int:
        return sum(self.counts)


@dataclass(frozen=True, slots=True)
class CarrierExportPreview:
    source_path: Path
    fingerprint: WorkbookFingerprint
    sheet_names: tuple[str, ...]
    months: tuple[tuple[int, int], ...]
    rows: tuple[CarrierExportRow, ...]
    carriers: tuple[CarrierCandidate, ...]
    unresolved_count: int


@dataclass(frozen=True, slots=True)
class CarrierExportResult:
    operation: ExcelOperation
    target_path: Path
    row_count: int
    total: Decimal
    carrier_name: str
    unresolved_count: int

    @property
    def message(self) -> str:
        return f"{self.carrier_name}: {self.row_count} dòng, tổng {self.total:,.0f} đ → {self.target_path}"


class CarrierExportService:
    def __init__(
        self,
        settings: Any | None = None,
        *,
        bk_path: str | Path | None = None,
        gateway: WorkbookGateway | None = None,
        run_repository: Any | None = None,
    ) -> None:
        self.bk_path = Path(bk_path or getattr(settings, "bk_workbook_path", ""))
        self.gateway = gateway or WorkbookGateway()
        self.months = MonthSheetService()
        self.run_repository = run_repository

    def source_sheet_candidates(self) -> list[SourceSheetCandidate]:
        path = ensure_supported_workbook(self.bk_path)
        if not path.is_file():
            raise CarrierExportError(f"Không tìm thấy file BK: {path}")
        book = self.gateway.load(path, read_only=True)
        try:
            return [SourceSheetCandidate(month=parsed[0], source_sheet=name)
                    for name in book.sheetnames
                    if (parsed := self.months.parse_target_sheet(name))]
        finally:
            book.close()

    def analyze(
        self,
        sheet_names: list[str] | tuple[str, ...],
        *,
        progress_callback: Callable[[str], None] | None = None,
    ) -> CarrierExportPreview:
        path = ensure_supported_workbook(self.bk_path)
        if not path.is_file():
            raise CarrierExportError(f"Không tìm thấy file BK: {path}")
        names = tuple(dict.fromkeys(sheet_names))
        if not names:
            raise CarrierExportError("Hãy chọn ít nhất một sheet tháng BK.")
        fingerprint = self.gateway.fingerprint(path)
        book = self.gateway.load(path)
        try:
            selected = []
            for name in names:
                parsed = self.months.parse_target_sheet(name)
                if parsed is None or name not in book:
                    raise CarrierExportError(f"Sheet tháng BK không hợp lệ: {name}")
                selected.append((name, *parsed))
            selected.sort(key=lambda entry: (entry[2], entry[1], book.sheetnames.index(entry[0])))
            normalization = NormalizationReport()
            for name, _, _ in selected:
                _ensure_bk_structure(book[name], normalization)
            rows: list[CarrierExportRow] = []
            names_by_key: dict[str, str] = {}
            counts: dict[str, Counter[str]] = {}
            for name, month, year in selected:
                if progress_callback:
                    progress_callback(f"Đang đọc {name}…")
                items, errors = _source_target_items(
                    book[name], normalization_issues=normalization.issues
                )
                for item in items["NAM"]:
                    if not any(_has_money(item.values.get(field)) for field in _FIELDS):
                        continue
                    carrier = normalize_carrier(item.carrier_value)
                    key = carrier.casefold() if carrier else None
                    error = errors.get(item.item_id)
                    try:
                        for field in _FIELDS:
                            _number(item.values.get(field))
                    except (ValueError, ArithmeticError):
                        error = error or "Khoản phí có số tiền không hợp lệ."
                    if any(len(values) > 1 for values in item.invoice_candidates.values()):
                        error = error or "Một khoản phí có nhiều số hóa đơn."
                    if error:
                        key = None
                    rows.append(CarrierExportRow(name, month, year, key, item, error))
                    if key:
                        names_by_key.setdefault(key, carrier)
                        counts.setdefault(key, Counter())[name] += 1
            self.gateway.assert_unchanged(path, fingerprint, label="File BK")
        finally:
            book.close()
        ordered_names = tuple(entry[0] for entry in selected)
        carriers = tuple(
            CarrierCandidate(key, names_by_key[key], tuple(counts[key][name] for name in ordered_names))
            for key in sorted(names_by_key, key=lambda k: names_by_key[k].casefold())
        )
        return CarrierExportPreview(
            path, fingerprint, ordered_names,
            tuple((entry[1], entry[2]) for entry in selected),
            tuple(rows), carriers, sum(row.carrier_key is None for row in rows),
        )

    @staticmethod
    def output_filename(preview: CarrierExportPreview, carrier_key: str) -> str:
        carrier = next((c for c in preview.carriers if c.key == carrier_key), None)
        if carrier is None:
            raise CarrierExportError("Bên vận tải đã chọn không có trong các sheet BK.")
        safe = _BAD_FILENAME.sub("_", carrier.name).rstrip(" .")
        if not safe:
            safe = "Ben_VT"
        cross_year = len({year for _, year in preview.months}) > 1
        suffix = "_".join(
            f"T{month:02d}_{year}" if cross_year else f"T{month:02d}"
            for month, year in preview.months
        )
        return f"{safe}_{suffix}.xlsx"

    def export(
        self,
        preview: CarrierExportPreview,
        carrier_key: str,
        output_dir: str | Path,
        *,
        overwrite: bool = False,
        progress_callback: Callable[[str], None] | None = None,
    ) -> CarrierExportResult:
        self.gateway.assert_unchanged(preview.source_path, preview.fingerprint, label="File BK")
        carrier = next((c for c in preview.carriers if c.key == carrier_key), None)
        if carrier is None:
            raise CarrierExportError("Bên vận tải đã chọn không còn trong danh sách.")
        rows = [row for row in preview.rows if row.carrier_key == carrier_key]
        if not rows:
            raise CarrierExportError("Bên vận tải đã chọn không có khoản phí trong Nam.")
        destination = Path(output_dir) / self.output_filename(preview, carrier_key)
        if destination.resolve() == preview.source_path.resolve():
            raise CarrierExportError("Không thể ghi đè file BK.")
        prior_fingerprint = workbook_fingerprint(destination) if destination.exists() else None
        if prior_fingerprint is not None and not overwrite:
            raise FileExistsError(f"File đã tồn tại: {destination}")
        destination.parent.mkdir(parents=True, exist_ok=True)
        run = None
        if self.run_repository is not None:
            run = self.run_repository.create_run(
                operation=ExcelOperation.CARRIER_EXPORT,
                status="APPLYING",
                source_path=preview.source_path,
                target_path=destination,
                source_fingerprint=preview.fingerprint,
                target_fingerprint_before=prior_fingerprint,
                sheet_name=", ".join(preview.sheet_names),
            )
        temp_path: Path | None = None
        try:
            if progress_callback:
                progress_callback(f"Đang tạo bảng kê cho {carrier.name}…")
            book = Workbook()
            sheet = book.active
            sheet.title = "Chi tiết phí trong Nam"
            sheet.append(_HEADERS)
            total = Decimal(0)
            for source in rows:
                item = source.item
                amounts = [_number(item.values.get(field)) for field in _FIELDS]
                amount_total = sum(amounts, Decimal(0))
                total += amount_total
                values: list[Any] = [source.sheet_name, item.sqt, item.container]
                for field, amount in zip(_FIELDS, amounts):
                    values.append(float(amount) if amount else None)
                    if field != "command_fee":
                        values.append(item.invoice_values.get(field))
                values += [
                    float(sum(amounts[i] for i in (1, 2, 3, 5))),
                    float(amounts[6] + amounts[8]),
                    float(amounts[7]),
                    float(amount_total),
                ]
                sheet.append(values)
                for index in (5, 7, 9, 11, 13, 16, 18, 20):
                    cell = sheet.cell(sheet.max_row, index)
                    if cell.value is not None:
                        cell.value = str(cell.value)
                        cell.number_format = "@"
            sheet.append(["TỔNG CỘNG", None, None] + [None] * 17 + [
                float(sum(_number(r.item.values.get(field)) for r in rows for field in ("loaded_lift", "empty_drop", "vs_do", "command_fee"))),
                float(sum(_number(r.item.values.get(field)) for r in rows for field in ("storage", "overweight"))),
                float(sum(_number(r.item.values.get("repair")) for r in rows)),
                float(total),
            ])
            self._style_sheet(sheet)
            handle, temp_name = tempfile.mkstemp(prefix=".carrier_export_", suffix=".xlsx", dir=destination.parent)
            os.close(handle)
            temp_path = Path(temp_name)
            book.save(temp_path)
            book.close()
            self._verify(temp_path, len(rows), total)
            self.gateway.assert_unchanged(preview.source_path, preview.fingerprint, label="File BK")
            if prior_fingerprint is None:
                if destination.exists():
                    raise FileExistsError(f"File đích vừa xuất hiện: {destination}")
            elif workbook_fingerprint(destination) != prior_fingerprint:
                raise WorkbookChangedError(label="File đích", path=destination)
            os.replace(temp_path, destination)
            temp_path = None
            result = CarrierExportResult(ExcelOperation.CARRIER_EXPORT, destination, len(rows), total, carrier.name, preview.unresolved_count)
            if run is not None:
                self.run_repository.finish_run(
                    run.id, status="SUCCEEDED", target_fingerprint_after=workbook_fingerprint(destination),
                    total_items=len(preview.rows), changed_items=len(rows), skipped_items=preview.unresolved_count,
                    item_outcomes={"carrier": carrier.name, "total": str(total)},
                )
            return result
        except Exception as exc:
            if run is not None:
                self.run_repository.finish_run(run.id, status="FAILED", error_message=str(exc))
            raise
        finally:
            if temp_path is not None:
                temp_path.unlink(missing_ok=True)

    @staticmethod
    def _style_sheet(sheet: Any) -> None:
        sheet.freeze_panes = "D2"
        sheet.auto_filter.ref = f"A1:X{sheet.max_row - 1}"
        sheet.print_options.horizontalCentered = True
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.print_title_rows = "1:1"
        sheet.print_area = f"A1:X{sheet.max_row}"
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="205781")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        sheet.row_dimensions[1].height = 42
        widths = {1: 16, 2: 9, 3: 19, 24: 18}
        for index in range(1, 25):
            sheet.column_dimensions[get_column_letter(index)].width = widths.get(index, 16)
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                if cell.column in (*_AMOUNT_COLUMNS, 21, 22, 23, 24):
                    cell.number_format = '#,##0.##;[Red](#,##0.##)'
        for cell in sheet[sheet.max_row]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DCEAF4")

    @staticmethod
    def _verify(path: Path, expected_rows: int, expected_total: Decimal) -> None:
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            if len(book.sheetnames) != 1:
                raise CarrierExportError("File xuất không có đúng một sheet.")
            sheet = book.active
            if sheet.max_row != expected_rows + 2 or tuple(cell.value for cell in sheet[1]) != _HEADERS:
                raise CarrierExportError("File xuất có sai số dòng hoặc sai cột.")
            calculated = sum((_number(sheet.cell(row, 24).value) for row in range(2, sheet.max_row)), Decimal(0))
            footer = _number(sheet.cell(sheet.max_row, 24).value)
            if abs(calculated - expected_total) > Decimal("0.01") or abs(footer - expected_total) > Decimal("0.01"):
                raise CarrierExportError("Tổng tiền trong file xuất không khớp dữ liệu BK.")
        finally:
            book.close()
