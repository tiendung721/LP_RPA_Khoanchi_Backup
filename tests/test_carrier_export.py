from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from PySide6.QtWidgets import QApplication, QDialog

from app.services.excel.carrier_export import (
    CarrierExportError, CarrierExportService, normalize_carrier, normalize_carriers,
)
from app.services.excel.payment_sync import BK_HEADER_ALIASES, SUMMARY_HEADERS
from app.services.excel.workbook import WorkbookChangedError, workbook_fingerprint
from app.database import Database
from app.repositories.excel_run_repository import ExcelRunRepository
from app.ui.carrier_export_dialog import CarrierAllocationDialog, CarrierRunSelectionDialog


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


def _allocate_mixed(service: CarrierExportService, preview):
    mixed = next(row for row in preview.allocation_rows if row.item.container == "EE")
    return service.apply_allocations(preview, {
        mixed.item.item_id: {
            "storage": [{"carrier_key": "anh quít", "amount": "100"}],
        },
    })


def test_normalize_only_explicit_month_suffix_and_distinct_names() -> None:
    assert normalize_carrier(" NHS   BK05 ") == "NHS"
    assert normalize_carrier("NHS BK T06 / NHS T06") == "NHS"
    assert normalize_carrier("A. Quít") == "A. Quít"
    assert normalize_carrier("Anh Quít") == "Anh Quít"
    assert normalize_carrier("A Quít 06") == "A Quít 06"
    assert normalize_carrier("NHS / Anh Quít") is None
    assert normalize_carrier("NHS BK07, NHS T07") == "NHS"
    assert normalize_carriers("NHS BK07, Anh Quít") == ("NHS", "Anh Quít")
    assert normalize_carriers("NHS + Anh Quít") == ("NHS", "Anh Quít")
    assert normalize_carriers("NHS; Anh Quít") == ("NHS", "Anh Quít")
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
    preview = _allocate_mixed(service, preview)
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
    preview = _allocate_mixed(service, preview)
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
        preview = _allocate_mixed(service, preview)
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


def test_comma_allocation_splits_money_and_invoices_per_run(tmp_path: Path) -> None:
    bk = tmp_path / "bk.xlsx"
    book = Workbook()
    book.remove(book.active)
    _sheet(book, "T07 26", [
        {"sqt": 1, "container": "AA", "carrier": "NHS BK07, Anh Quít",
         "loaded_lift": 100, "invoice_loaded_lift": "00017", "empty_drop": 60},
        {"sqt": 2, "container": "BB", "carrier": "NHS BK07, NHS T07",
         "storage": 30},
    ])
    book.save(bk)
    book.close()
    database = Database(tmp_path / "history.db")
    try:
        service = CarrierExportService(
            bk_path=bk, run_repository=ExcelRunRepository(database)
        )
        preview = service.analyze(["T07 26"])
        assert preview.run_id is not None
        assert len(preview.allocation_rows) == 1
        assert preview.unresolved_count == 1
        mixed = preview.allocation_rows[0]
        with pytest.raises(CarrierExportError, match="chưa phân bổ"):
            service.export(preview, "nhs", tmp_path)
        choices = {mixed.item.item_id: {
            "loaded_lift": [
                {"carrier_key": "nhs", "amount": "40", "invoice": "00017-A"},
                {"carrier_key": "anh quít", "amount": "60", "invoice": "00017-B"},
            ],
            "empty_drop": [{"carrier_key": "anh quít", "amount": "60"}],
        }}
        with pytest.raises(CarrierExportError, match="tổng phân bổ"):
            bad = {mixed.item.item_id: {**choices[mixed.item.item_id],
                   "loaded_lift": [
                       {"carrier_key": "nhs", "amount": "40"},
                       {"carrier_key": "anh quít", "amount": "50"},
                   ]}}
            service.apply_allocations(preview, bad)
        with pytest.raises(CarrierExportError, match="hóa đơn"):
            missing_invoice = {mixed.item.item_id: {**choices[mixed.item.item_id],
                   "loaded_lift": [
                       {"carrier_key": "nhs", "amount": "40"},
                       {"carrier_key": "anh quít", "amount": "60"},
                   ]}}
            service.apply_allocations(preview, missing_invoice)
        resolved = service.apply_allocations(preview, choices)
        assert resolved.unresolved_count == 0
        assert {carrier.name: carrier.total for carrier in resolved.carriers} == {
            "NHS": 2, "Anh Quít": 1,
        }
        nhs_result = service.export(resolved, "nhs", tmp_path)
        assert nhs_result.total == 70
        exported = load_workbook(nhs_result.target_path, read_only=True)
        try:
            assert exported.active["F2"].value == 40
            assert exported.active["G2"].value == "00017-A"
            assert exported.active["X4"].value == 70
        finally:
            exported.close()
        anh_result = service.export(resolved, "anh quít", tmp_path)
        assert anh_result.total == 120
        exported = load_workbook(anh_result.target_path, read_only=True)
        try:
            assert exported.active["F2"].value == 60
            assert exported.active["G2"].value == "00017-B"
            assert exported.active["H2"].value == 60
            assert exported.active["X3"].value == 120
        finally:
            exported.close()
        resumed = service.analyze(["T07 26"], resume_run_id=preview.run_id)
        assert resumed.unresolved_count == 0
        assert resumed.allocation_choices == resolved.allocation_choices
        assert resumed.selected_carrier_key == "anh quít"
        fresh = service.analyze(["T07 26"])
        assert fresh.run_id != preview.run_id
        assert fresh.unresolved_count == 1
        book = load_workbook(bk)
        book["T07 26"]["A2"] = "note"  # unrelated column leaves the fee signature intact
        book.save(bk)
        book.close()
        still_valid = service.analyze(["T07 26"], resume_run_id=preview.run_id)
        assert still_valid.unresolved_count == 0
        book = load_workbook(bk)
        amount_column = next(
            cell.column for cell in book["T07 26"][1]
            if cell.value == BK_HEADER_ALIASES["loaded_lift"][0]
        )
        book["T07 26"].cell(2, amount_column, 110)
        book.save(bk)
        book.close()
        changed = service.analyze(["T07 26"], resume_run_id=preview.run_id)
        assert changed.unresolved_count == 1
        assert not changed.allocation_choices
    finally:
        database.close()


def test_allocation_dialog_saves_one_run_and_run_picker_reopens_it(tmp_path: Path) -> None:
    app = QApplication.instance() or QApplication([])
    bk = tmp_path / "bk.xlsx"
    _bk(bk)
    database = Database(tmp_path / "history.db")
    try:
        service = CarrierExportService(
            bk_path=bk, run_repository=ExcelRunRepository(database)
        )
        preview = service.analyze(["T08 26"])
        dialog = CarrierAllocationDialog(preview, service)
        assert dialog.row_list.count() == 1
        combo, _inputs = dialog._controls["storage"]
        combo.setCurrentIndex(combo.findData("anh quít"))
        dialog._accept_allocations()
        assert dialog.result() == QDialog.DialogCode.Accepted
        assert dialog.resolved_preview is not None
        assert dialog.resolved_preview.unresolved_count == 1
        runs = service.saved_runs()
        picker = CarrierRunSelectionDialog(runs)
        picker._reopen()
        assert picker.resume_run_id == preview.run_id
        picker.close()
        dialog.close()
    finally:
        database.close()
