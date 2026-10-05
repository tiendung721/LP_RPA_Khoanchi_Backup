from __future__ import annotations

import json
from dataclasses import replace
from pathlib import Path

from openpyxl import Workbook, load_workbook
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QPushButton

from app.config import AppSettings
from app.database import Database
from app.repositories.excel_draft_repository import ExcelDraftRepository
from app.repositories.rpa_tracking_repository import RpaTrackingRepository
from app.rpa_expense import (
    RPA_STATUS_IMPORTED,
    RPA_STATUS_NOT_IMPORTED,
    RpaExpenseService,
    RpaExpenseStatusService,
    RpaChoiceService,
)
from app.rpa_expense.launcher import RpaExpenseBatLauncher
from app.rpa_expense.service import STATUS_HEADER, SUMMARY_HEADERS
from app.services.excel.models import (
    ExcelRunStatus,
    FieldWriteOutcome,
    ItemWriteOutcome,
    OutcomeStatus,
    PostingResult,
)
from app.services.excel.posting import ExpensePostingService
from app.ui.rpa_expense_controller import RpaExpenseController
from app.ui.rpa_expense_dialog import RpaLatestDataDialog, RpaSqtSelectionDialog


def _build_bk(path: Path, *, with_status: bool = True) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "T07 26"
    sheet["A1"] = "SQT"
    sheet["B1"] = "Số HĐ"
    summary_start = 3
    for offset, header in enumerate(SUMMARY_HEADERS):
        sheet.cell(1, summary_start + offset).value = header
    if with_status:
        sheet.cell(1, summary_start + len(SUMMARY_HEADERS)).value = STATUS_HEADER

    # Hai dòng cùng SQT 101 để kiểm tra tổng hợp nhiều dòng.
    rows = (
        (101, (100, 20, 30, 40, 50, 0, 5, 0, 60), None),
        (101, (200, 10, 0, 5, 25, 0, 15, 0, 40), RPA_STATUS_IMPORTED),
        (102, (300, 30, 10, 0, 20, 0, 0, 0, 50), RPA_STATUS_IMPORTED),
    )
    for row_number, (sqt, values, status) in enumerate(rows, 2):
        sheet.cell(row_number, 1).value = sqt
        sheet.cell(row_number, 2).value = f"HD-{sqt}-{row_number}"
        sheet.cell(row_number, summary_start).value = f"=A{row_number}"
        for offset, value in enumerate(values, 1):
            sheet.cell(row_number, summary_start + offset).value = value
        if with_status:
            sheet.cell(
                row_number,
                summary_start + len(SUMMARY_HEADERS),
            ).value = status
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def _settings(tmp_path: Path, bk: Path) -> AppSettings:
    return AppSettings(
        data_root=tmp_path / "runtime",
        output_dir=tmp_path / "Output",
        bk_workbook_path=str(bk),
    )


def _add_second_sheet(path: Path, *, duplicate_sqt: bool = False) -> None:
    workbook = load_workbook(path)
    try:
        sheet = workbook.copy_worksheet(workbook["T07 26"])
        sheet.title = "T08 26"
        if not duplicate_sqt:
            for row, sqt in ((2, 901), (3, 901), (4, 902)):
                sheet.cell(row, 1).value = sqt
                sheet.cell(row, 2).value = f"HD-{sqt}-{row}"
        workbook.save(path)
    finally:
        workbook.close()


def test_analyze_groups_rows_and_imported_items_remain_runnable(
    tmp_path: Path,
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    service = RpaExpenseService(_settings(tmp_path, bk))

    assert [item.sheet_name for item in service.sheet_candidates()] == ["T07 26"]
    plan = service.analyze_sheet("T07 26")
    first, second = plan.items

    assert first.sqt == "101"
    assert first.source_rows == (2, 3)
    assert first.invoice_numbers == ("HD-101-2", "HD-101-3")
    assert first.status == RPA_STATUS_NOT_IMPORTED
    assert first.amounts.cuoc_bo_dong_hang == 300
    assert first.amounts.nang_ha_dong_hang == 30
    assert first.amounts.luu_cont_qua_tai == 100
    assert first.amounts.sua_chua_cont == 20
    assert first.can_run

    assert second.status == RPA_STATUS_IMPORTED
    assert second.can_run


def test_one_pad_selection_updates_sqt_in_two_bk_sheets(
    qtbot, tmp_path: Path
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    _add_second_sheet(bk)
    settings = _settings(tmp_path, bk)
    service = RpaExpenseService(settings)
    plan = service.analyze_all_sheets()

    assert plan.sheet_name == ""
    assert {(item.sqt, item.sheet_name) for item in plan.items} == {
        ("101", "T07 26"), ("102", "T07 26"),
        ("901", "T08 26"), ("902", "T08 26"),
    }
    dialog = RpaSqtSelectionDialog(plan)
    qtbot.addWidget(dialog)
    assert dialog.table.columnCount() == len(dialog.MULTI_COLUMNS)
    assert dialog.table.horizontalHeaderItem(2).text() == "Sheet BK"
    assert dialog.table.horizontalHeaderItem(3).text() == "HĐ theo loại phí"

    prepared = service.prepare_selection(plan, ["101", "901"])
    payload = json.loads(prepared.selection_path.read_text(encoding="utf-8"))
    assert payload["sheet_name"] is None
    assert [(item["sqt"], item["sheet_name"]) for item in payload["items"]] == [
        ("101", "T07 26"), ("901", "T08 26")
    ]
    status = RpaExpenseStatusService(backup_dir=settings.paths.excel_backup_dir)
    status.mark_imported(prepared.selection_path, "901")
    workbook = load_workbook(bk)
    try:
        status_column = 3 + len(SUMMARY_HEADERS)
        assert workbook["T08 26"].cell(2, status_column).value == RPA_STATUS_IMPORTED
        assert workbook["T07 26"].cell(2, status_column).value is None
    finally:
        workbook.close()
    status.mark_imported(prepared.selection_path, "101")
    workbook = load_workbook(bk)
    try:
        assert workbook["T07 26"].cell(2, status_column).value == RPA_STATUS_IMPORTED
        assert workbook["T07 26"].cell(3, status_column).value == RPA_STATUS_IMPORTED
    finally:
        workbook.close()


def test_duplicate_sqt_across_sheets_cannot_be_sent_to_pad(tmp_path: Path) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    _add_second_sheet(bk, duplicate_sqt=True)
    plan = RpaExpenseService(_settings(tmp_path, bk)).analyze_all_sheets()
    assert plan.runnable_count == 0
    assert all("SQT trùng" in item.validation_message for item in plan.items)


def test_status_helper_still_reads_legacy_single_sheet_json(tmp_path: Path) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    settings = _settings(tmp_path, bk)
    service = RpaExpenseService(settings)
    prepared = service.prepare_selection(service.analyze_sheet("T07 26"), ["101"])
    payload = json.loads(prepared.selection_path.read_text(encoding="utf-8"))
    payload["items"][0].pop("sheet_name")
    prepared.selection_path.write_text(json.dumps(payload), encoding="utf-8")
    result = RpaExpenseStatusService(
        backup_dir=settings.paths.excel_backup_dir
    ).mark_imported(prepared.selection_path, "101")
    assert result["sheet_name"] == "T07 26"


def test_analyze_collects_distinct_invoices_from_multiple_bk_columns(
    tmp_path: Path,
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "T07 26"
    sheet["A1"] = "SQT"
    sheet["B1"] = "Số HĐ"
    sheet["C1"] = "Hóa đơn cước biển"
    for offset, header in enumerate(SUMMARY_HEADERS, 4):
        sheet.cell(1, offset).value = header
    for row, invoices in ((2, ("HD-01", "SEA-02")), (3, ("hd-01", "SEA-03"))):
        sheet.cell(row, 1).value = 101
        sheet.cell(row, 2).value = invoices[0]
        sheet.cell(row, 3).value = invoices[1]
        sheet.cell(row, 4).value = f"=A{row}"
    bk.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(bk)
    workbook.close()

    plan = RpaExpenseService(_settings(tmp_path, bk)).analyze_sheet("T07 26")
    assert plan.items[0].invoice_numbers == ("HD-01", "SEA-02", "SEA-03")


def test_fee_invoices_keep_their_fee_and_source_row_without_changing_pad_json(
    qtbot, tmp_path: Path
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "T08 26"
    headers = (
        "SQT", "Cước bộ đóng hàng", "Số HĐ", "Nâng vỏ", "Hóa đơn",
        "Hạ Hàng", "Hóa đơn", "Cước biển", "Hóa đơn cước biển",
        "Lưu cont", "HD", "Quá tải", "HD", "GHI CHÚ",
        "SỬA CHỮA", "HD",
    )
    for column, header in enumerate((*headers, *SUMMARY_HEADERS), 1):
        sheet.cell(1, column).value = header
    sheet.cell(2, 1).value = 900
    for column, value in {
        2: 100, 3: "MB-1", 4: 50, 5: "NV-1", 6: 20, 7: "HH-1",
        8: 70, 9: "SEA-1", 10: 10, 11: "LC-1", 12: 5, 13: "QT-1",
        15: 8, 16: "SC-1",
    }.items():
        sheet.cell(2, column).value = value
    summary_start = len(headers) + 1
    for offset, value in enumerate(("=A2", 100, 70, 70, 0, 0, 10, 8, 5, 15)):
        sheet.cell(2, summary_start + offset).value = value
    bk.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(bk)
    workbook.close()

    service = RpaExpenseService(_settings(tmp_path, bk))
    plan = service.analyze_all_sheets()
    item = plan.items[0]
    assert {(entry.fee_label, entry.invoice_number, entry.source_row)
            for entry in item.fee_entries} == {
        ("Cước bộ đóng hàng", "MB-1", 2),
        ("Nâng vỏ", "NV-1", 2),
        ("Hạ hàng", "HH-1", 2),
        ("Cước biển", "SEA-1", 2),
        ("Lưu cont", "LC-1", 2),
        ("Quá tải", "QT-1", 2),
        ("Sửa chữa", "SC-1", 2),
    }
    dialog = RpaSqtSelectionDialog(plan)
    qtbot.addWidget(dialog)
    assert "Nâng vỏ: NV-1" in dialog.table.item(0, 3).text()
    assert "Hạ hàng: HH-1" in dialog.table.item(0, 3).text()
    assert any(
        dialog.detail_table.item(row, 2).text() == "HH-1"
        and dialog.detail_table.item(row, 3).text() == "2"
        for row in range(dialog.detail_table.rowCount())
    )
    dialog.sqt_search.setText("QT-1")
    assert dialog.table.item(0, 3).text().startswith("Quá tải: QT-1")
    assert any(
        dialog.detail_table.item(row, 2).text() == "QT-1"
        and dialog.detail_table.item(row, 2).background().color().name()
        == "#fff4cc"
        for row in range(dialog.detail_table.rowCount())
    )
    payload_item = service.prepare_selection(plan, ["900"]).payload["items"][0]
    assert "fee_entries" not in payload_item
    assert "invoice_numbers" not in payload_item


def test_sheet_dropdown_keeps_cross_month_checks_and_invoice_search(
    qtbot, tmp_path: Path
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    _add_second_sheet(bk)
    dialog = RpaSqtSelectionDialog(
        RpaExpenseService(_settings(tmp_path, bk)).analyze_all_sheets()
    )
    qtbot.addWidget(dialog)
    assert dialog.sheet_filter.itemText(0) == "Tất cả (4)"
    assert dialog.sheet_filter.itemData(1) == "T08 26"

    dialog.sheet_filter.setCurrentIndex(dialog.sheet_filter.findData("T07 26"))
    visible = [
        row for row in range(dialog.table.rowCount())
        if not dialog.table.isRowHidden(row)
    ]
    assert {dialog.table.item(row, 1).text() for row in visible} == {"101", "102"}
    first = next(row for row in visible if dialog.table.item(row, 1).text() == "101")
    dialog.table.item(first, 0).setCheckState(Qt.CheckState.Checked)

    dialog.sheet_filter.setCurrentIndex(dialog.sheet_filter.findData("T08 26"))
    visible = [
        row for row in range(dialog.table.rowCount())
        if not dialog.table.isRowHidden(row)
    ]
    assert {dialog.table.item(row, 1).text() for row in visible} == {"901", "902"}
    second = next(row for row in visible if dialog.table.item(row, 1).text() == "901")
    dialog.table.item(second, 0).setCheckState(Qt.CheckState.Checked)
    assert set(dialog.selected_sqt) == {"101", "901"}

    dialog.sqt_search.setText("HD-901-2")
    assert [
        dialog.table.item(row, 1).text()
        for row in range(dialog.table.rowCount())
        if not dialog.table.isRowHidden(row)
    ] == ["901"]
    assert "SQT 901" in dialog.detail_title.text()
    assert any(
        dialog.detail_table.item(row, 2).text() == "HD-901-2"
        for row in range(dialog.detail_table.rowCount())
    )
    dialog.sqt_search.clear()
    assert set(dialog.selected_sqt) == {"101", "901"}


def test_multi_sheet_dialog_keeps_sqt_list_visible_when_resized(
    qtbot, tmp_path: Path
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    _add_second_sheet(bk)
    dialog = RpaSqtSelectionDialog(
        RpaExpenseService(_settings(tmp_path, bk)).analyze_all_sheets()
    )
    qtbot.addWidget(dialog)
    dialog.show()
    dialog.resize(1100, 750)
    qtbot.wait(20)

    assert not dialog.detail_panel.isVisible()
    assert dialog.table.height() >= 300
    assert dialog.table.rowCount() == 4
    dialog.table.item(0, 0).setCheckState(Qt.CheckState.Checked)
    checked_before = set(dialog.selected_sqt)

    dialog.findChild(QPushButton, "rpaDetailToggle").click()
    assert dialog._detail_popup is not None
    assert dialog._detail_popup.isVisible()
    assert dialog._popup_table.rowCount() == dialog.detail_table.rowCount()

    dialog.resize(1400, 760)
    qtbot.wait(20)
    assert dialog.detail_panel.isVisible()
    assert not dialog._detail_popup.isVisible()
    assert set(dialog.selected_sqt) == checked_before


def test_seal_invoice_is_labeled_as_outside_pad(tmp_path: Path) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    workbook = load_workbook(bk)
    try:
        sheet = workbook["T07 26"]
        sheet.insert_cols(3)
        sheet.cell(1, 3).value = "Số HĐ Seal"
        sheet.cell(2, 3).value = "SEAL-1"
        workbook.save(bk)
    finally:
        workbook.close()
    service = RpaExpenseService(_settings(tmp_path, bk))
    item = service.analyze_sheet("T07 26").items[0]
    assert any(
        entry.invoice_number == "SEAL-1"
        and entry.fee_label == "Phí Seal (không gửi PAD)"
        for entry in item.fee_entries
    )
    assert "SEAL-1" not in str(item.to_payload())


def test_prepare_json_and_mark_all_source_rows_after_success(
    tmp_path: Path,
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    settings = _settings(tmp_path, bk)
    service = RpaExpenseService(settings)
    plan = service.analyze_sheet("T07 26")

    prepared = service.prepare_selection(plan, ["101", "102"])
    payload = json.loads(prepared.selection_path.read_text(encoding="utf-8"))

    assert payload["version"] == 1
    assert payload["operation"] == "NHAP_KHOAN_CHI_BK"
    assert payload["items"][0]["source_rows"] == [2, 3]
    assert payload["items"][0]["status_before"] == RPA_STATUS_NOT_IMPORTED
    assert payload["status_callback"]["when"] == "AFTER_WEB_SAVE_SUCCESS"
    assert payload["status_callback"]["arguments"][-1] == "{sqt}"

    status = RpaExpenseStatusService(
        backup_dir=settings.paths.excel_backup_dir
    )
    first_result = status.mark_imported(prepared.selection_path, "101")
    second_result = status.mark_imported(prepared.selection_path, "102")

    workbook = load_workbook(bk, data_only=False)
    try:
        sheet = workbook["T07 26"]
        status_column = 3 + len(SUMMARY_HEADERS)
        assert sheet.cell(2, status_column).value == RPA_STATUS_IMPORTED
        assert sheet.cell(3, status_column).value == RPA_STATUS_IMPORTED
        assert sheet.cell(4, status_column).value == RPA_STATUS_IMPORTED
    finally:
        workbook.close()
    assert first_result["backup_path"] == second_result["backup_path"]
    assert len(list(settings.paths.excel_backup_dir.glob("*.xlsx"))) == 1


def test_status_column_is_created_when_missing(tmp_path: Path) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk, with_status=False)
    settings = _settings(tmp_path, bk)
    service = RpaExpenseService(settings)
    prepared = service.prepare_selection(
        service.analyze_sheet("T07 26"),
        ["101"],
    )

    RpaExpenseStatusService(
        backup_dir=settings.paths.excel_backup_dir
    ).mark_imported(prepared.selection_path, "101")

    workbook = load_workbook(bk, data_only=False)
    try:
        sheet = workbook["T07 26"]
        status_column = 3 + len(SUMMARY_HEADERS)
        assert sheet.cell(1, status_column).value == STATUS_HEADER
        assert sheet.cell(2, status_column).value == RPA_STATUS_IMPORTED
        assert sheet.cell(3, status_column).value == RPA_STATUS_IMPORTED
    finally:
        workbook.close()


def test_dialog_keeps_imported_sqt_selectable(qtbot, tmp_path: Path) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    plan = RpaExpenseService(_settings(tmp_path, bk)).analyze_sheet("T07 26")
    dialog = RpaSqtSelectionDialog(plan)
    qtbot.addWidget(dialog)

    imported_check = dialog.table.item(1, 0)
    assert imported_check.flags() & Qt.ItemFlag.ItemIsUserCheckable
    imported_check.setCheckState(Qt.CheckState.Checked)
    assert dialog.selected_sqt == ["102"]


def test_dialog_filters_searches_and_selects_all_rows(
    qtbot, tmp_path: Path
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    plan = RpaExpenseService(_settings(tmp_path, bk)).analyze_sheet("T07 26")
    dialog = RpaSqtSelectionDialog(plan, initial_selected_sqt=["102"])
    qtbot.addWidget(dialog)

    assert dialog.filter_buttons[None].text() == "Tất cả (2)"
    assert (
        dialog.filter_buttons[RPA_STATUS_NOT_IMPORTED].text()
        == "Chưa nhập (1)"
    )
    assert dialog.filter_buttons[RPA_STATUS_IMPORTED].text() == "Đã nhập (1)"
    assert dialog.table.item(0, 2).text() == "HD-101-2, HD-101-3"
    assert dialog.table.item(0, 3).text() == "● Chưa nhập"
    assert dialog.table.item(1, 3).text() == "✓ Đã nhập"
    assert "1 nhập lại" in dialog.selection_summary.text()

    dialog.filter_buttons[RPA_STATUS_NOT_IMPORTED].click()
    visible_sqt = [
        dialog.table.item(row, 1).text()
        for row in range(dialog.table.rowCount())
        if not dialog.table.isRowHidden(row)
    ]
    assert visible_sqt == ["101"]

    dialog.findChild(QPushButton, "selectAllRpaSqtButton").click()
    assert dialog.selected_sqt == ["101", "102"]
    assert "1 chưa nhập, 1 nhập lại" in dialog.selection_summary.text()

    dialog.findChild(QPushButton, "clearAllRpaSqtButton").click()
    assert dialog.selected_sqt == ["102"]

    dialog.filter_buttons[None].click()
    dialog.findChild(QPushButton, "clearAllRpaSqtButton").click()
    assert dialog.selected_sqt == []

    dialog.sqt_search.setText("02")
    visible_sqt = [
        dialog.table.item(row, 1).text()
        for row in range(dialog.table.rowCount())
        if not dialog.table.isRowHidden(row)
    ]
    assert visible_sqt == ["102"]
    dialog.findChild(QPushButton, "selectAllRpaSqtButton").click()
    assert dialog.selected_sqt == ["102"]
    dialog.sqt_search.setText("hd-101-3")
    visible_sqt = [
        dialog.table.item(row, 1).text()
        for row in range(dialog.table.rowCount())
        if not dialog.table.isRowHidden(row)
    ]
    assert visible_sqt == ["101"]
    dialog.findChild(QPushButton, "clearAllRpaSqtButton").click()
    assert dialog.selected_sqt == ["102"]
    dialog.sqt_search.clear()
    assert all(
        not dialog.table.isRowHidden(row)
        for row in range(dialog.table.rowCount())
    )


def test_dialog_removes_total_column_and_sorts_sqt_by_numeric_value(
    qtbot, tmp_path: Path
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    plan = RpaExpenseService(_settings(tmp_path, bk)).analyze_sheet("T07 26")
    dialog = RpaSqtSelectionDialog(plan)
    qtbot.addWidget(dialog)

    assert "Tổng" not in dialog.COLUMNS
    assert dialog.table.columnCount() == len(dialog.COLUMNS) == 14
    assert dialog.sort_combo.findData("total_asc") == -1
    assert (
        dialog.table.horizontalScrollBarPolicy()
        == Qt.ScrollBarPolicy.ScrollBarAlwaysOn
    )

    dialog.sort_combo.setCurrentIndex(dialog.sort_combo.findData("sqt_desc"))
    assert [
        dialog.table.item(row, 1).text()
        for row in range(dialog.table.rowCount())
    ] == ["102", "101"]


def test_zero_amount_sqt_can_be_selected_singly_or_in_bulk(
    qtbot,
    tmp_path: Path,
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    workbook = load_workbook(bk, data_only=False)
    try:
        sheet = workbook["T07 26"]
        for row_number in range(2, 5):
            for column in range(4, 13):
                sheet.cell(row_number, column).value = 0
        workbook.save(bk)
    finally:
        workbook.close()

    service = RpaExpenseService(_settings(tmp_path, bk))
    plan = service.analyze_sheet("T07 26")
    assert plan.runnable_count == 2
    assert all(item.amounts.total == 0 for item in plan.items)
    assert all(item.can_run for item in plan.items)
    assert all(
        item.validation_message == "Tất cả khoản tiền đều bằng 0."
        for item in plan.items
    )

    dialog = RpaSqtSelectionDialog(plan)
    qtbot.addWidget(dialog)
    first_check = dialog.table.item(0, 0)
    second_check = dialog.table.item(1, 0)
    assert first_check.flags() & Qt.ItemFlag.ItemIsUserCheckable
    assert second_check.flags() & Qt.ItemFlag.ItemIsUserCheckable

    first_check.setCheckState(Qt.CheckState.Checked)
    assert dialog.selected_sqt == ["101"]

    second_check.setCheckState(Qt.CheckState.Checked)
    assert dialog.selected_sqt == ["101", "102"]

    prepared = service.prepare_selection(plan, dialog.selected_sqt)
    assert prepared.item_count == 2
    assert [
        sum(item["amounts"].values()) for item in prepared.payload["items"]
    ] == [0, 0]


def test_launcher_passes_the_fixed_json_path_to_bat(
    monkeypatch,
    tmp_path: Path,
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    bat = tmp_path / "run_pad.bat"
    bat.write_text("@echo off\r\n", encoding="utf-8")
    settings = _settings(tmp_path, bk)
    settings.rpa_expense_bat_path = str(bat)
    service = RpaExpenseService(settings)
    prepared = service.prepare_selection(
        service.analyze_sheet("T07 26"),
        ["101"],
    )
    captured: dict[str, object] = {}

    def start_detached(
        command: str,
        arguments: list[str],
        working_directory: str,
    ) -> tuple[bool, int]:
        captured.update(
            command=command,
            arguments=arguments,
            working_directory=working_directory,
        )
        return True, 4321

    monkeypatch.setattr(
        "app.rpa_expense.launcher.QProcess.startDetached",
        start_detached,
    )
    result = RpaExpenseBatLauncher(settings).launch(prepared)

    assert result.process_id == 4321
    assert captured["arguments"][-1] == str(prepared.selection_path)
    assert prepared.selection_path.name == "rpa_input_selection.json"
    assert captured["working_directory"] == str(bat.parent.resolve())


def test_latest_rpa_snapshot_changes_only_after_recording_a_successful_launch(
    tmp_path: Path,
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    service = RpaExpenseService(_settings(tmp_path, bk))
    plan = service.analyze_sheet("T07 26")
    first = service.prepare_selection(plan, ["101"])

    assert service.load_latest_launched() is None
    service.record_launched(first)
    persisted = service.load_latest_launched()
    assert persisted is not None
    assert persisted["run_id"] == first.run_id
    assert [item["sqt"] for item in persisted["items"]] == ["101"]
    assert persisted["launched_at"]

    service.prepare_selection(plan, ["102"])
    assert service.load_latest_launched()["run_id"] == first.run_id


def test_multi_sheet_launch_history_tracks_each_sheet(tmp_path: Path) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    _add_second_sheet(bk)
    database = Database(tmp_path / "app.db")
    try:
        tracking = RpaTrackingRepository(database)
        service = RpaExpenseService(
            _settings(tmp_path, bk), tracking_repository=tracking
        )
        prepared = service.prepare_selection(
            service.analyze_all_sheets(), ["101", "901"]
        )
        service.record_launched(prepared)
        assert tracking.snapshot(bk, "T07 26").latest_pad_sqt == ("101",)
        assert tracking.snapshot(bk, "T08 26").latest_pad_sqt == ("901",)
    finally:
        database.close()


def test_latest_rpa_dialog_displays_every_sent_sqt(qtbot) -> None:
    payload = {
        "run_id": "run-1",
        "sheet_name": "T07 26",
        "launched_at": "2026-08-19T10:00:00+07:00",
        "items": [
            {
                "sqt": value,
                "source_rows": [row],
                "status_before": "Chưa nhập",
                "amounts": {"cuoc_bien": amount},
            }
            for value, row, amount in (("101", 2, 100), ("102", 3, 200))
        ],
    }
    dialog = RpaLatestDataDialog(payload)
    qtbot.addWidget(dialog)

    assert dialog.table.rowCount() == 2
    assert dialog.table.item(0, 0).text() == "101"
    assert dialog.table.item(1, 0).text() == "102"


def test_latest_rpa_dialog_shows_source_sheet_for_multi_sheet_run(qtbot) -> None:
    dialog = RpaLatestDataDialog({
        "run_id": "run-1",
        "sheet_name": None,
        "items": [
            {"sqt": "101", "sheet_name": "T07 26", "amounts": {}},
            {"sqt": "901", "sheet_name": "T08 26", "amounts": {}},
        ],
    })
    qtbot.addWidget(dialog)
    assert dialog.table.columnCount() == len(dialog.COLUMNS) + 1
    assert dialog.table.item(0, len(dialog.COLUMNS)).text() == "T07 26"
    assert dialog.table.item(1, len(dialog.COLUMNS)).text() == "T08 26"


def test_rpa_choices_restore_imported_sqt_and_reject_changed_amounts(
    tmp_path: Path,
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    service = RpaExpenseService(_settings(tmp_path, bk))
    database = Database(tmp_path / "app.db")
    choices = RpaChoiceService(ExcelDraftRepository(database))
    plan = service.analyze_sheet("T07 26")

    choices.save_selection(plan, ["101", "102"])
    restored = choices.restore_selection(plan)
    assert restored.selected_sqt == ("101", "102")
    assert restored.restored_count == 2

    workbook = load_workbook(bk, data_only=False)
    try:
        workbook["T07 26"].cell(2, 4).value = 999
        workbook.save(bk)
    finally:
        workbook.close()
    changed = choices.restore_selection(service.analyze_sheet("T07 26"))
    assert changed.selected_sqt == ("102",)
    assert changed.skipped_sqt == ("101",)
    database.close()


def test_rpa_choice_dialog_prefills_and_clears_remembered_selection(
    qtbot, tmp_path: Path
) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    plan = RpaExpenseService(_settings(tmp_path, bk)).analyze_sheet("T07 26")
    cleared: list[bool] = []
    dialog = RpaSqtSelectionDialog(
        plan,
        initial_selected_sqt=["101", "102"],
        restore_info={"found": True, "saved_count": 2, "restored_count": 2},
        clear_saved_callback=lambda: cleared.append(True) or True,
    )
    qtbot.addWidget(dialog)

    assert dialog.selected_sqt == ["101", "102"]
    dialog.findChild(QPushButton, "clearRememberedRpaSelectionButton").click()
    assert cleared == [True]
    assert dialog.selected_sqt == []


def test_rpa_choices_are_isolated_by_workbook_and_sheet(tmp_path: Path) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    service = RpaExpenseService(_settings(tmp_path, bk))
    database = Database(tmp_path / "app.db")
    choices = RpaChoiceService(ExcelDraftRepository(database))
    plan = service.analyze_sheet("T07 26")
    other_sheet_plan = type(plan)(
        bk_path=plan.bk_path,
        sheet_name="T08 26",
        fingerprint=plan.fingerprint,
        items=plan.items,
    )

    choices.save_sheet(bk, "T07 26")
    choices.save_selection(plan, ["101"])
    choices.save_selection(other_sheet_plan, ["102"])

    candidates = service.sheet_candidates()
    assert choices.restore_sheet(bk, candidates) == "T07 26"
    assert choices.restore_selection(plan).selected_sqt == ("101",)
    assert choices.restore_selection(other_sheet_plan).selected_sqt == ("102",)
    database.close()


def test_rpa_controller_keeps_choices_when_launch_fails(qtbot, tmp_path: Path) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    service = RpaExpenseService(_settings(tmp_path, bk))
    database = Database(tmp_path / "app.db")
    choices = RpaChoiceService(ExcelDraftRepository(database))
    plan = service.analyze_sheet("T07 26")

    class FailingLauncher:
        @staticmethod
        def launch(_prepared):
            raise RuntimeError("BAT không khởi chạy")

    controller = RpaExpenseController(service, FailingLauncher(), choices)
    failures: list[Exception] = []
    controller.failed.connect(failures.append)
    try:
        controller.launch(plan, ["102"])
        qtbot.waitUntil(lambda: len(failures) == 1)

        restored = choices.restore_selection(plan)
        assert restored.selected_sqt == ("102",)
        assert restored.status == "FAILED"
    finally:
        controller.shutdown()
        database.close()


def test_rpa_tracking_replaces_latest_bk_across_workbook_and_latest_pad(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "app.db")
    tracking = RpaTrackingRepository(database)
    bk = tmp_path / "BK.xlsx"
    try:
        tracking.record_bk_run(
            bk,
            {"T07 26": ["101", "102"], "T08 26": ["201"]},
        )
        assert tracking.snapshot(bk, "T08 26").latest_bk_revisions == {"201": 1}
        tracking.record_bk_run(bk, {"T07 26": ["102", "103"]})
        snapshot = tracking.snapshot(bk, "T07 26")

        assert snapshot.latest_bk_revisions == {"102": 2, "103": 1}
        assert tracking.bk_revision(bk, "T07 26", "101") == 1

        tracking.record_bk_run(bk, {"T08 26": ["201"]})
        assert tracking.snapshot(bk, "T07 26").latest_bk_revisions == {}
        assert tracking.snapshot(bk, "T08 26").latest_bk_revisions == {"201": 2}
        tracking.record_bk_run(bk, {})
        assert tracking.snapshot(bk, "T08 26").latest_bk_revisions == {}

        tracking.record_latest_pad(
            bk,
            "T07 26",
            ["90", "101"],
            run_id="run-1",
        )
        tracking.record_latest_pad(
            bk,
            "T07 26",
            ["102"],
            run_id="run-2",
        )
        assert tracking.snapshot(bk, "T07 26").latest_pad_sqt == ("102",)
    finally:
        database.close()


def test_pad_success_rejects_stale_bk_revision(tmp_path: Path) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    settings = _settings(tmp_path, bk)
    database = Database(tmp_path / "app.db")
    tracking = RpaTrackingRepository(database)
    try:
        tracking.record_bk_run(bk, {"T07 26": ["101"]})
        service = RpaExpenseService(settings, tracking_repository=tracking)
        plan = service.analyze_sheet("T07 26")
        prepared = service.prepare_selection(plan, ["101"])
        payload = json.loads(prepared.selection_path.read_text(encoding="utf-8"))
        assert payload["items"][0]["bk_revision"] == 1

        # BK lại thay đổi trong lúc payload cũ đang được PAD xử lý.
        tracking.record_bk_run(bk, {"T07 26": ["101"]})
        result = RpaExpenseStatusService(
            backup_dir=settings.paths.excel_backup_dir,
            tracking_repository=tracking,
        ).mark_imported(prepared.selection_path, "101")

        assert result["stale_selection"] is True
        assert tracking.snapshot(bk, "T07 26").latest_bk_revisions == {"101": 2}
        workbook = load_workbook(bk, data_only=False)
        try:
            status_column = 3 + len(SUMMARY_HEADERS)
            assert workbook["T07 26"].cell(2, status_column).value is None
        finally:
            workbook.close()
    finally:
        database.close()


def test_pad_success_keeps_latest_bk_and_records_latest_launch(tmp_path: Path) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    settings = _settings(tmp_path, bk)
    database = Database(tmp_path / "app.db")
    tracking = RpaTrackingRepository(database)
    try:
        tracking.record_bk_run(bk, {"T07 26": ["101"]})
        service = RpaExpenseService(settings, tracking_repository=tracking)
        prepared = service.prepare_selection(
            service.analyze_sheet("T07 26"),
            ["101"],
        )
        service.record_launched(prepared)
        assert tracking.snapshot(bk, "T07 26").latest_pad_sqt == ("101",)

        result = RpaExpenseStatusService(
            backup_dir=settings.paths.excel_backup_dir,
            tracking_repository=tracking,
        ).mark_imported(prepared.selection_path, "101")

        assert result["success"] is True
        assert tracking.snapshot(bk, "T07 26").latest_bk_revisions == {"101": 1}
    finally:
        database.close()


def test_dialog_selects_and_filters_rpa_groups(qtbot, tmp_path: Path) -> None:
    bk = tmp_path / "Output" / "BK.xlsx"
    _build_bk(bk)
    base_plan = RpaExpenseService(_settings(tmp_path, bk)).analyze_sheet("T07 26")
    plan = replace(
        base_plan,
        latest_bk_revisions=(("101", 1),),
        latest_pad_sqt=("102",),
    )
    dialog = RpaSqtSelectionDialog(plan)
    qtbot.addWidget(dialog)

    assert set(dialog.selected_sqt) == {"101"}
    assert dialog.group_filter_buttons["bk"].isChecked()
    visible = {
        dialog.table.item(row, 1).text()
        for row in range(dialog.table.rowCount())
        if not dialog.table.isRowHidden(row)
    }
    assert visible == {"101"}
    assert dialog.table.item(0, 4).text() == "Vừa ghi BK"

    dialog.sqt_search.setText("HD-102-4")
    visible = {
        dialog.table.item(row, 1).text()
        for row in range(dialog.table.rowCount())
        if not dialog.table.isRowHidden(row)
    }
    assert visible == {"102"}
    dialog.sqt_search.clear()

    assert dialog.findChild(QPushButton, "selectLatestRpaGroupButton") is None
    assert dialog.findChild(QPushButton, "selectBothRpaGroupsButton") is None
    dialog.findChild(QPushButton, "clearAllRpaSqtButton").click()
    assert dialog.selected_sqt == []
    dialog.findChild(QPushButton, "selectAllRpaSqtButton").click()
    assert dialog.selected_sqt == ["101"]
    dialog.group_filter_buttons["all"].click()
    dialog.findChild(QPushButton, "selectAllRpaSqtButton").click()
    assert set(dialog.selected_sqt) == {"101", "102"}


def test_posting_result_tracks_only_written_amounts_in_latest_bk(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")
    tracking = RpaTrackingRepository(database)
    bk = tmp_path / "BK.xlsx"
    service = ExpensePostingService(
        provider=object(),
        bk_path=bk,
        rpa_tracking_repository=tracking,
    )
    try:
        service._finish_result(
            PostingResult(
                status=ExcelRunStatus.SUCCEEDED,
                target_path=bk,
                item_outcomes=[
                    ItemWriteOutcome(
                        item_id="posting:1",
                        sqt=101,
                        target_sheet="T07 26",
                        fields=[
                            FieldWriteOutcome(
                                field_name="Số tiền",
                                status=OutcomeStatus.WRITTEN,
                            )
                        ],
                    ),
                    ItemWriteOutcome(
                        item_id="posting:2",
                        sqt=102,
                        target_sheet="T07 26",
                        fields=[
                            FieldWriteOutcome(
                                field_name="Số tiền",
                                status=OutcomeStatus.USER_KEPT,
                            )
                        ],
                    ),
                ],
            )
        )

        assert tracking.snapshot(bk, "T07 26").latest_bk_revisions == {"101": 1}
    finally:
        database.close()


def test_posting_amount_change_resets_existing_rpa_status() -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.cell(1, 5).value = STATUS_HEADER
    sheet.cell(2, 5).value = RPA_STATUS_IMPORTED

    ExpensePostingService._reset_rpa_statuses(
        sheet,
        [{"amount_write": True, "target_row": 2}],
    )

    assert sheet.cell(2, 5).value == RPA_STATUS_NOT_IMPORTED
    workbook.close()
