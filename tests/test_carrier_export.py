from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from app.services.excel.carrier_export import CarrierExportService, normalize_carrier
from app.services.excel.payment_sync import BK_HEADER_ALIASES, SUMMARY_HEADERS
from app.services.excel.workbook import WorkbookChangedError, workbook_fingerprint
from app.database import Database
from app.repositories.excel_run_repository import ExcelRunRepository


FIELDS = (
    "sea_freight", "north_freight", "empty_lift", "loaded_drop",
    "loaded_lift", "empty_drop", "south_freight", "storage", "overweight",
    "vs_do", "seal", "command_fee", "repair",
)


def _sheet(book: Workbook, name: str, rows: list[dict[str, object]]) -> None:
    sheet = book.create_sheet(name)
    sheet.cell(1, 1, "Ngày")  # Headers need not begin in column A.
    columns = {"sqt": 2, "container": 3}
    for field in ("sqt", "container"):
        sheet.cell(1, columns[field], BK_HEADER_ALIASES[field][0])
    next_column = 4
    invoice_columns = {}
    for field in FIELDS:
        columns[field] = next_column
        sheet.cell(1, next_column, BK_HEADER_ALIASES[field][0])
        next_column += 1
        if field not in {"sea_freight", "north_freight", "south_freight", "command_fee"}:
            invoice_columns[field] = next_column
            sheet.cell(1, next_column, "Số HĐ Seal" if field == "seal" else "Số HĐ")
            next_column += 1
    columns["carrier_nam"] = next_column
    sheet.cell(1, next_column, BK_HEADER_ALIASES["carrier_nam"][0])
    for column, header in enumerate(SUMMARY_HEADERS, next_column + 1):
        sheet.cell(1, column, header)
    for index, data in enumerate(rows, 2):
        sheet.cell(index, columns["sqt"], data.get("sqt", index))
        sheet.cell(index, columns["container"], data.get("container", f"CONT{index}"))
        sheet.cell(index, columns["carrier_nam"], data.get("carrier"))
        for field in FIELDS:
            if field in data:
                sheet.cell(index, columns[field], data[field])
            if f"invoice_{field}" in data and field in invoice_columns:
                sheet.cell(index, invoice_columns[field], data[f"invoice_{field}"])


def _bk(path: Path) -> None:
    book = Workbook()
    book.remove(book.active)
    _sheet(book, "T08 26", [
        {"sqt": 1, "container": "AA", "carrier": "NHS BK08", "empty_lift": "=100+50", "invoice_empty_lift": "00017"},
        {"sqt": 1, "container": "AA", "carrier": "NHS BK T08", "loaded_lift": 200},
        {"sqt": 2, "container": "BB", "carrier": "A. Quít", "repair": 500},
        {"sqt": 3, "container": "CC", "carrier": "Anh Quít", "repair": 700},
        {"sqt": 4, "container": "DD", "carrier": None, "overweight": 100},
        {"sqt": 5, "container": "EE", "carrier": "NHS / Anh Quít", "storage": 100},
        {"sqt": 6, "container": "FF", "carrier": "NHS BK08", "loaded_drop": 999},
    ])
    _sheet(book, "T09 26", [
        {"sqt": 1, "container": "AA", "carrier": "NHS BK09", "seal": 10, "vs_do": 20,
         "command_fee": 30, "storage": 40, "repair": 50, "overweight": 60, "empty_drop": 70},
    ])
    _sheet(book, "T10 26", [])
    book.save(path)
    book.close()


def test_normalize_only_explicit_month_suffix_and_distinct_names() -> None:
    assert normalize_carrier(" NHS   BK05 ") == "NHS"
    assert normalize_carrier("NHS BK T06 / NHS T06") == "NHS"
    assert normalize_carrier("A. Quít") == "A. Quít"
    assert normalize_carrier("Anh Quít") == "Anh Quít"
    assert normalize_carrier("A Quít 06") == "A Quít 06"
    assert normalize_carrier("NHS / Anh Quít") is None
    assert normalize_carrier("0") is None


def test_export_multiple_months_one_sheet_and_preserves_bk(tmp_path: Path) -> None:
    bk = tmp_path / "bk.xlsx"
    _bk(bk)
    before = workbook_fingerprint(bk)
    service = CarrierExportService(bk_path=bk)
    preview = service.analyze(["T10 26", "T09 26", "T08 26"])
    assert preview.sheet_names == ("T08 26", "T09 26", "T10 26")
    assert preview.unresolved_count == 2
    assert {candidate.name for candidate in preview.carriers} == {"NHS", "A. Quít", "Anh Quít"}
    nhs = next(candidate for candidate in preview.carriers if candidate.name == "NHS")
    assert nhs.counts == (1, 1, 0)
    assert service.output_filename(preview, nhs.key) == "NHS_T08_T09_T10.xlsx"
    result = service.export(preview, nhs.key, tmp_path)
    assert result.row_count == 2
    assert result.total == 630
    assert workbook_fingerprint(bk) == before
    book = load_workbook(result.target_path)
    try:
        assert len(book.sheetnames) == 1
        sheet = book.active
        assert sheet.max_row == 4
        assert [sheet.cell(row, 1).value for row in (2, 3)] == ["T08 26", "T09 26"]
        assert [sheet.cell(row, 3).value for row in (2, 3)] == ["AA", "AA"]
        assert sheet["D2"].value == 150
        assert sheet["E2"].value == "00017"
        assert sheet["E2"].data_type == "s"
        assert sheet["F2"].value == 200
        assert sheet["L3"].value == 10
        assert sheet["M1"].value == "Số HĐ Seal"
        assert sheet["U3"].value == 120
        assert sheet["V3"].value == 100
        assert sheet["W3"].value == 50
        assert sheet["X4"].value == 630
        assert sheet["X3"].value == 280
    finally:
        book.close()


def test_existing_target_and_changed_bk_never_overwrite(tmp_path: Path) -> None:
    bk = tmp_path / "bk.xlsx"
    _bk(bk)
    service = CarrierExportService(bk_path=bk)
    preview = service.analyze(["T08 26"])
    key = next(candidate.key for candidate in preview.carriers if candidate.name == "NHS")
    result = service.export(preview, key, tmp_path)
    exported = result.target_path.read_bytes()
    with pytest.raises(FileExistsError):
        service.export(preview, key, tmp_path)
    assert result.target_path.read_bytes() == exported
    book = load_workbook(bk)
    book["T08 26"]["A2"] = "changed"
    book.save(bk)
    book.close()
    with pytest.raises(WorkbookChangedError):
        service.export(preview, key, tmp_path, overwrite=True)
    assert result.target_path.read_bytes() == exported


def test_overwrite_history_and_cross_year_name(tmp_path: Path) -> None:
    bk = tmp_path / "bk.xlsx"
    _bk(bk)
    book = load_workbook(bk)
    _sheet(book, "T01 27", [{"sqt": 1, "container": "AA", "carrier": "NHS BK01", "seal": 5}])
    book.save(bk)
    book.close()
    before = workbook_fingerprint(bk)
    database = Database(tmp_path / "history.db")
    try:
        repository = ExcelRunRepository(database)
        service = CarrierExportService(bk_path=bk, run_repository=repository)
        preview = service.analyze(["T08 26", "T01 27"])
        key = next(candidate.key for candidate in preview.carriers if candidate.name == "NHS")
        assert service.output_filename(preview, key) == "NHS_T08_2026_T01_2027.xlsx"
        target = tmp_path / service.output_filename(preview, key)
        target.write_bytes(b"old file")
        result = service.export(preview, key, tmp_path, overwrite=True)
        assert result.row_count == 2
        assert result.total == 355
        record = repository.get_latest(operation="CARRIER_EXPORT", statuses=("SUCCEEDED",))
        assert record is not None
        assert record.changed_items == 2
        assert record.target_path == target
        assert workbook_fingerprint(bk) == before
    finally:
        database.close()
