from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest
from PySide6.QtCore import QItemSelectionModel, Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFrame,
    QHeaderView,
    QMessageBox,
    QPushButton,
    QTableWidget,
)

import app.ui.main_window as main_window_module
from app.database import Database
from app.repositories.excel_run_repository import ExcelRunRepository
from app.ui.excel_dialogs import (
    ConflictResolutionDialog,
    ExcelOutcomeDialog,
    ManualRowPickerDialog,
    MonthSelectionDialog,
    PaymentNewRowsDialog,
    RepostSelectionDialog,
)
from app.services.excel.models import (
    FieldWriteOutcome,
    ItemWriteOutcome,
    OutcomeStatus,
)
from app.ui.excel_task_controller import ExcelTaskController
from app.ui.excel_summary_dialogs import (
    ExcelCompletionDialog,
    ExcelConfirmationDialog,
    daily_confirmation_summary,
    payment_completion_summary,
    payment_confirmation_summary,
    posting_completion_summary,
    posting_confirmation_summary,
)
from app.ui.main_window import MainWindow
from app.ui.settings_page import SettingsPage
from app.ui.workflow_page import WorkflowPage


def test_step_three_has_three_primary_actions_and_statuses(qtbot) -> None:
    page = WorkflowPage()
    qtbot.addWidget(page)

    buttons = page.step3_card.findChildren(QPushButton)
    primary_buttons = [button for button in buttons if button.property("primary")]

    assert primary_buttons == [
        page.sync_daily_button,
        page.post_expenses_button,
        page.sync_payment_button,
    ]
    assert page.sync_daily_button.text() == "Đồng bộ"
    assert page.post_expenses_button.text() == "Nhập vào BK"
    assert page.sync_payment_button.text() == "Đồng bộ"
    assert page.sync_status_label.text() == "Đồng bộ gần nhất: —"
    assert page.posting_status_label.text() == "Nhập khoản chi gần nhất: —"
    assert (
        page.payment_sync_status_label.text()
        == "Đồng bộ BK → Thanh toán gần nhất: —"
    )


def test_excel_confirmation_dialog_renders_cards_notice_table_and_explicit_action(
    qtbot,
) -> None:
    summary = daily_confirmation_summary(
        SimpleNamespace(),
        candidates=[
            SimpleNamespace(
                source_sheet="Tháng 7",
                target_sheet="T07 26",
                update_count=5,
                new_row_count=2,
                unchanged_count=10,
                target_only_count=1,
                invalid_count=3,
            )
        ],
    )
    dialog = ExcelConfirmationDialog(
        summary,
        confirm_label="Đồng bộ 7 thay đổi vào BK",
    )
    qtbot.addWidget(dialog)

    assert dialog.windowTitle() == "Xác nhận đồng bộ Hàng ngày → BK"
    assert len(dialog.findChildren(QFrame, "excelSummaryMetric")) == 4
    assert len(dialog.findChildren(QFrame, "excelSummaryNotice")) == 2
    table = dialog.findChild(QTableWidget, "excelSummaryTable")
    assert table is not None
    assert table.item(0, 0).text() == "Tháng 7"
    assert table.item(0, 1).text() == "T07 26"
    assert dialog.confirm_button.text() == "Đồng bộ 7 thay đổi vào BK"
    assert not dialog.confirm_button.isDefault()

    qtbot.mouseClick(dialog.confirm_button, Qt.MouseButton.LeftButton)
    assert dialog.result() == QDialog.DialogCode.Accepted


def test_excel_completion_dialog_returns_open_and_detail_actions(qtbot) -> None:
    summary = posting_completion_summary(
        SimpleNamespace(
            status="SUCCEEDED",
            sheet_name="T07 26",
            posted_source_items=2,
            already_existing_items=4,
            skipped_source_items=1,
            item_outcomes=[],
        )
    )
    open_dialog = ExcelCompletionDialog(
        summary,
        open_label="Mở file BK",
        details_available=True,
    )
    qtbot.addWidget(open_dialog)
    qtbot.mouseClick(open_dialog.open_button, Qt.MouseButton.LeftButton)
    assert open_dialog.selected_action == ExcelCompletionDialog.OPEN_TARGET

    detail_dialog = ExcelCompletionDialog(
        summary,
        open_label="Mở file BK",
        details_available=True,
    )
    qtbot.addWidget(detail_dialog)
    assert detail_dialog.details_button is not None
    qtbot.mouseClick(detail_dialog.details_button, Qt.MouseButton.LeftButton)
    assert detail_dialog.selected_action == ExcelCompletionDialog.VIEW_DETAILS


def test_outcome_dialog_flattens_fields_and_filters_errors(qtbot) -> None:
    outcomes = [
        ItemWriteOutcome(
            item_id="one",
            container="CONT1",
            status=OutcomeStatus.PARTIAL,
            fields=[
                FieldWriteOutcome(
                    field_name="Số tiền",
                    status=OutcomeStatus.WRITTEN,
                    reason="Đã ghi",
                ),
                FieldWriteOutcome(
                    field_name="Số HĐ",
                    status=OutcomeStatus.USER_KEPT,
                    reason="User giữ",
                ),
            ],
        ),
        ItemWriteOutcome(
            item_id="bad",
            status=OutcomeStatus.INVALID_SOURCE,
            fields=[
                FieldWriteOutcome(
                    field_name="Dòng nguồn",
                    status=OutcomeStatus.INVALID_SOURCE,
                    reason="Thiếu SQT",
                )
            ],
        ),
    ]
    dialog = ExcelOutcomeDialog(outcomes)
    qtbot.addWidget(dialog)

    assert dialog.table.rowCount() == 3
    dialog.filter_combo.setCurrentIndex(4)
    assert dialog.table.rowCount() == 1
    assert dialog.table.item(0, 10).text() == "Thiếu SQT"


def test_outcome_dialog_displays_friendly_vietnamese_statuses(qtbot) -> None:
    expected_labels = {
        OutcomeStatus.WRITTEN: "Đã cập nhật",
        OutcomeStatus.PARTIAL: "Chỉ cập nhật một phần",
        OutcomeStatus.UNCHANGED: "Không cần thay đổi",
        OutcomeStatus.USER_KEPT: "Giữ nguyên theo lựa chọn",
        OutcomeStatus.USER_SKIPPED: "Bỏ qua theo lựa chọn",
        OutcomeStatus.INVALID_SOURCE: "Dữ liệu nguồn không hợp lệ",
        OutcomeStatus.FAILED: "Xử lý không thành công",
    }
    outcomes = [
        ItemWriteOutcome(
            item_id=status.value,
            fields=[
                FieldWriteOutcome(
                    field_name=status.value,
                    status=status,
                    reason=status.value,
                )
            ],
        )
        for status in expected_labels
    ]
    dialog = ExcelOutcomeDialog(outcomes)
    qtbot.addWidget(dialog)

    labels_by_code = {
        dialog.table.item(row, 10).text(): dialog.table.item(row, 9).text()
        for row in range(dialog.table.rowCount())
    }
    assert labels_by_code == {
        status.value: label for status, label in expected_labels.items()
    }
    assert [
        dialog.filter_combo.itemText(index)
        for index in range(dialog.filter_combo.count())
    ] == [
        "Tất cả",
        "Đã cập nhật",
        "Không cần thay đổi",
        "Theo lựa chọn người dùng",
        "Cần kiểm tra",
    ]


def test_outcome_dialog_translates_payment_field_codes_and_formats_amounts(
    qtbot,
) -> None:
    dialog = ExcelOutcomeDialog(
        [
            ItemWriteOutcome(
                item_id="payment-row",
                fields=[
                    FieldWriteOutcome(
                        field_name="Dòng thanh toán",
                        source_value={
                            "loaded_drop": 1_252_800,
                            "storage": 1_209_600,
                        },
                        target_value_before={"loaded_drop": 1_000_000},
                        target_value_after={
                            "loaded_drop": 1_252_800,
                            "storage": 1_209_600,
                        },
                        status=OutcomeStatus.WRITTEN,
                    )
                ],
            )
        ]
    )
    qtbot.addWidget(dialog)

    expected = "Hạ hàng: 1.252.800; Lưu cont: 1.209.600"
    assert dialog.table.item(0, 4).text() == expected
    assert dialog.table.item(0, 7).text() == "Hạ hàng: 1.000.000"
    assert dialog.table.item(0, 8).text() == expected
    assert "loaded_drop" not in dialog.table.item(0, 4).text()
    header = dialog.table.horizontalHeader()
    for column in (4, 8):
        assert (
            header.sectionResizeMode(column)
            is QHeaderView.ResizeMode.Interactive
        )
        assert dialog.table.columnWidth(column) == 240
    assert (
        dialog.table.horizontalScrollMode()
        is dialog.table.ScrollMode.ScrollPerPixel
    )
    assert (
        dialog.table.horizontalScrollBarPolicy()
        is Qt.ScrollBarPolicy.ScrollBarAlwaysOn
    )
    assert (
        header.sectionResizeMode(10)
        is QHeaderView.ResizeMode.Interactive
    )
    assert dialog.table.columnWidth(10) == 360
    dialog.show()
    qtbot.waitUntil(
        lambda: dialog.table.horizontalScrollBar().maximum() > 0
    )


def test_main_window_restores_latest_excel_details_after_restart(
    monkeypatch,
    qtbot,
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "app_state.db")
    repository = ExcelRunRepository(database)
    run = repository.create_run(operation="DAILY_SYNC")
    repository.finish_run(
        run.id,
        status="SUCCEEDED",
        sheet_name="T07 26",
        item_outcomes=[
            {
                "item_id": "daily-row-2",
                "status": "WRITTEN",
                "fields": [{"field_name": "Dòng đồng bộ A–K/P"}],
            }
        ],
    )
    captured: dict[str, Any] = {}

    class Dialog:
        def __init__(self, outcomes: Any, *, parent: Any = None) -> None:
            captured["outcomes"] = list(outcomes)
            captured["parent"] = parent

        @staticmethod
        def exec() -> None:
            captured["opened"] = True

    monkeypatch.setattr(main_window_module, "ExcelOutcomeDialog", Dialog)
    window = MainWindow(
        settings={},
        excel_run_repository=repository,
        start_watcher=False,
    )
    qtbot.addWidget(window)

    assert window.workflow_page.view_daily_sync_button.isEnabled()
    window.workflow_page.view_daily_sync_button.click()
    assert captured["opened"] is True
    assert captured["outcomes"][0]["item_id"] == "daily-row-2"
    database.close()


def test_main_window_restores_latest_rpa_payload_after_restart(
    monkeypatch,
    qtbot,
) -> None:
    payload = {
        "operation": "NHAP_KHOAN_CHI_BK",
        "run_id": "run-1",
        "sheet_name": "T07 26",
        "items": [{"sqt": "101", "amounts": {}}],
    }
    controller = SimpleNamespace(load_latest_launched=lambda: payload)
    captured: dict[str, Any] = {}

    class Dialog:
        def __init__(self, value: Any, parent: Any = None) -> None:
            captured["payload"] = value
            captured["parent"] = parent

        @staticmethod
        def exec() -> None:
            captured["opened"] = True

    monkeypatch.setattr(main_window_module, "RpaLatestDataDialog", Dialog)
    window = MainWindow(
        settings={},
        rpa_expense_controller=controller,
        start_watcher=False,
    )
    qtbot.addWidget(window)

    assert window.workflow_page.view_rpa_expense_button.isEnabled()
    window.workflow_page.view_rpa_expense_button.click()
    assert captured["opened"] is True
    assert captured["payload"]["run_id"] == "run-1"


def test_step_three_locks_all_excel_actions_while_running(qtbot) -> None:
    page = WorkflowPage()
    qtbot.addWidget(page)

    page.set_excel_running("sync", "Đang đọc file Hàng ngày")

    assert not page.sync_daily_button.isEnabled()
    assert not page.post_expenses_button.isEnabled()
    assert not page.sync_payment_button.isEnabled()
    assert page.sync_daily_button.text() == "Đang đồng bộ…"
    assert "Đang đọc file Hàng ngày" in page.sync_status_label.text()

    page.set_excel_idle("sync")

    assert page.sync_daily_button.isEnabled()
    assert page.post_expenses_button.isEnabled()
    assert page.sync_payment_button.isEnabled()
    assert page.sync_daily_button.text() == "Đồng bộ"


def test_sync_button_requires_source_sheet_before_starting_analysis(
    monkeypatch,
) -> None:
    candidates = [
        SimpleNamespace(month=7, sheet_name="Tháng 7"),
        SimpleNamespace(month=8, sheet_name="Tháng 8"),
    ]

    class Service:
        def source_sheet_candidates(self) -> list[Any]:
            return candidates

    class Tasks:
        daily_sync_service = Service()

        def __init__(self) -> None:
            self.sync_calls: list[dict[str, Any]] = []

        def start_sync(self, **kwargs: Any) -> None:
            self.sync_calls.append(kwargs)

    captured: dict[str, Any] = {}

    class Dialog:
        def __init__(
            self,
            dialog_candidates: list[Any],
            _parent: Any,
            **kwargs: Any,
        ) -> None:
            captured["candidates"] = dialog_candidates
            captured.update(kwargs)
            self.selected_sheet_names = ["Tháng 8"]

        def exec(self) -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

    monkeypatch.setattr(main_window_module, "MonthSelectionDialog", Dialog)
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_context=None,
        _missing_excel_configuration=lambda **_kwargs: False,
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow.start_daily_sync(owner)

    assert captured["candidates"] == candidates
    assert captured["preselect_first"] is False
    assert captured["show_recommendations"] is False
    assert captured["multi_select"] is True
    assert tasks.sync_calls == [{"source_sheet_names": ["Tháng 8"]}]


def test_sync_analysis_always_shows_full_sheet_confirmation(
    monkeypatch,
) -> None:
    candidate = SimpleNamespace(
        month=7,
        target_sheet="T07 26",
        update_count=5,
        new_row_count=2,
        unchanged_count=10,
        target_only_count=1,
        invalid_count=3,
    )
    plan = SimpleNamespace(
        operation="sync",
        conflicts=[],
        selected_month=7,
        selected_sheet="T07 26",
        month_candidates=[candidate],
    )

    class Tasks:
        def __init__(self) -> None:
            self.apply_calls: list[Any] = []
            self.cancel_calls = 0

        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "sync"

        def apply_plan(
            self,
            value: Any,
            resolutions: Any,
            *,
            operation: str,
        ) -> None:
            self.apply_calls.append((value, resolutions, operation))

        def cancel_waiting(self) -> None:
            self.cancel_calls += 1

    class Confirmation:
        summary: Any = None
        confirm_label = ""

        def __init__(self, summary: Any, *, confirm_label: str, parent: Any) -> None:
            Confirmation.summary = summary
            Confirmation.confirm_label = confirm_label

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

    monkeypatch.setattr(main_window_module, "ExcelConfirmationDialog", Confirmation)
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_operation="sync",
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow._excel_analysis_ready(owner, plan)

    assert len(tasks.apply_calls) == 1
    assert tasks.cancel_calls == 0
    assert Confirmation.summary.title == "Xác nhận đồng bộ Hàng ngày → BK"
    assert [metric.value for metric in Confirmation.summary.metrics] == [5, 2, 10, 3]
    assert Confirmation.summary.tables[0].rows[0][1:] == (
        "T07 26", 5, 2, 10, 1, 3
    )
    assert Confirmation.confirm_label == "Đồng bộ 7 thay đổi vào BK"


def test_sync_without_changes_applies_without_redundant_confirmation(
    monkeypatch,
) -> None:
    candidate = SimpleNamespace(
        source_sheet="Tháng 7",
        target_sheet="T07 26",
        update_count=0,
        new_row_count=0,
        unchanged_count=12,
        target_only_count=2,
        invalid_count=0,
    )
    plan = SimpleNamespace(
        operation="sync",
        conflicts=[],
        selected_month=7,
        selected_sheet="T07 26",
        month_candidates=[candidate],
    )

    class Tasks:
        def __init__(self) -> None:
            self.apply_calls: list[Any] = []

        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "sync"

        def apply_plan(self, value: Any, resolutions: Any, *, operation: str) -> None:
            self.apply_calls.append((value, resolutions, operation))

    class UnexpectedConfirmation:
        def __init__(self, *_args: Any, **_kwargs: Any) -> None:
            raise AssertionError("Không được hỏi xác nhận khi không có thay đổi.")

    monkeypatch.setattr(
        main_window_module,
        "ExcelConfirmationDialog",
        UnexpectedConfirmation,
    )
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_operation="sync",
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow._excel_analysis_ready(owner, plan)

    assert tasks.apply_calls == [(plan, {}, "sync")]


def test_bang_ke_posting_skips_global_month_and_applies_multi_sheet_plan(
    monkeypatch,
) -> None:
    plan = SimpleNamespace(
        operation="posting",
        source_kind="BANG_KE",
        batch_id=42,
        conflicts=[],
        selected_sheet=None,
        sheet_candidates=[
            SimpleNamespace(target_sheet="T01 26"),
            SimpleNamespace(target_sheet="T04 26"),
        ],
        target_sheets={"T04 26", "T01 26"},
        previously_posted_items=[],
        repost_selection_done=False,
        confirmation_required=True,
        confirmation_done=False,
        original_source_count=5,
        reconciliation_source_count=0,
        items=[
            SimpleNamespace(
                source_indices=[0],
                status="PLANNED",
                action="OVERWRITE",
                cell_state=SimpleNamespace(kind="EMPTY"),
            ),
            SimpleNamespace(
                source_indices=[1],
                status="USER_SKIPPED",
                action="KEEP_EXISTING",
                cell_state=SimpleNamespace(kind="NUMBER"),
            ),
        ],
    )

    class Tasks:
        def __init__(self) -> None:
            self.apply_calls: list[Any] = []
            self.cancel_calls = 0

        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "posting"

        @staticmethod
        def start_posting(**_kwargs: Any) -> None:
            raise AssertionError("BANG_KE không được phân tích lại theo một sheet chung.")

        def apply_plan(
            self,
            value: Any,
            resolutions: Any,
            *,
            operation: str,
        ) -> None:
            self.apply_calls.append((value, resolutions, operation))

        def cancel_waiting(self) -> None:
            self.cancel_calls += 1

    class UnexpectedMonthDialog:
        def __init__(self, *_args: Any, **_kwargs: Any) -> None:
            raise AssertionError("BANG_KE không được hỏi một tháng chung.")

    class Confirmation:
        summary: Any = None
        confirm_label = ""

        def __init__(self, summary: Any, *, confirm_label: str, parent: Any) -> None:
            Confirmation.summary = summary
            Confirmation.confirm_label = confirm_label

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

    monkeypatch.setattr(
        main_window_module, "MonthSelectionDialog", UnexpectedMonthDialog
    )
    monkeypatch.setattr(main_window_module, "ExcelConfirmationDialog", Confirmation)
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_operation="posting",
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow._excel_analysis_ready(owner, plan)

    assert tasks.apply_calls == [(plan, {}, "posting")]
    assert tasks.cancel_calls == 0
    assert plan.confirmation_done is True
    assert Confirmation.summary.title == "Xác nhận nhập khoản chi vào BK"
    assert Confirmation.summary.subtitle.endswith("2 sheet.")
    assert [metric.value for metric in Confirmation.summary.metrics] == [1, 1, 0, 1]
    assert Confirmation.confirm_label == "Nhập 1 khoản vào BK"


def test_posting_confirmation_summarizes_file_effects_in_plain_language() -> None:
    plan = SimpleNamespace(
        selected_sheet="T07 26",
        target_sheets=set(),
        items=[
            SimpleNamespace(
                source_indices=[0],
                status="PLANNED",
                action="OVERWRITE",
                cell_state=SimpleNamespace(kind="EMPTY"),
            ),
            SimpleNamespace(
                source_indices=[1],
                status="PLANNED",
                action="OVERWRITE",
                cell_state=SimpleNamespace(kind="NUMBER"),
            ),
            SimpleNamespace(
                source_indices=[2],
                status="USER_SKIPPED",
                action="KEEP_EXISTING",
                cell_state=SimpleNamespace(kind="NUMBER"),
            ),
            SimpleNamespace(
                source_indices=[3, 4],
                status="USER_SKIPPED",
                action="SKIP",
                cell_state=None,
            ),
            SimpleNamespace(
                source_indices=[5],
                status="ALREADY_EXISTS",
                action=None,
                cell_state=SimpleNamespace(kind="SAME_VALUE"),
            ),
        ],
    )

    summary = posting_confirmation_summary(plan)

    assert summary.subtitle == "2 ô BK sẽ thay đổi trên 1 sheet."
    assert [metric.value for metric in summary.metrics] == [2, 2, 1, 3]
    assert summary.tables[0].rows == (("T07 26", 1, 1, 1, 1, 2),)
    assert summary.notices[0].tone == "warning"
    assert "1 khoản sẽ thay thế" in summary.notices[0].text


def test_regular_posting_still_reanalyzes_the_selected_month(monkeypatch) -> None:
    candidates = [
        SimpleNamespace(target_sheet="T06 26"),
        SimpleNamespace(target_sheet="T07 26"),
    ]
    plan = SimpleNamespace(
        operation="posting",
        source_kind="ASSISTANT",
        batch_id=7,
        conflicts=[],
        selected_sheet=None,
        sheet_candidates=candidates,
    )

    class Tasks:
        def __init__(self) -> None:
            self.start_calls: list[dict[str, Any]] = []
            self.cancel_calls = 0

        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "posting"

        def start_posting(self, **kwargs: Any) -> None:
            self.start_calls.append(kwargs)

        def cancel_waiting(self) -> None:
            self.cancel_calls += 1

    class Dialog:
        selected_sheet_name = "T07 26"

        def __init__(
            self,
            dialog_candidates: list[Any],
            _parent: Any,
            **_kwargs: Any,
        ) -> None:
            assert dialog_candidates == candidates

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

    monkeypatch.setattr(main_window_module, "MonthSelectionDialog", Dialog)
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_context="workflow",
        _excel_operation="posting",
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow._excel_analysis_ready(owner, plan)

    assert tasks.cancel_calls == 1
    assert tasks.start_calls == [{"batch_id": 7, "sheet_name": "T07 26"}]


def test_bang_ke_repost_selection_reanalyzes_without_a_global_month(
    monkeypatch,
) -> None:
    plan = SimpleNamespace(
        operation="posting",
        source_kind="BANG_KE",
        batch_id=42,
        conflicts=[],
        selected_sheet=None,
        sheet_candidates=[SimpleNamespace(target_sheet="T01 26")],
        previously_posted_items=[SimpleNamespace(source_item_index=3)],
        repost_selection_done=False,
    )

    class Tasks:
        def __init__(self) -> None:
            self.start_calls: list[dict[str, Any]] = []
            self.cancel_calls = 0

        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "posting"

        def start_posting(self, **kwargs: Any) -> None:
            self.start_calls.append(kwargs)

        def cancel_waiting(self) -> None:
            self.cancel_calls += 1

    class Dialog:
        selected_source_indices = {3}

        def __init__(self, items: list[Any], _parent: Any) -> None:
            assert items == plan.previously_posted_items

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

    class UnexpectedMonthDialog:
        def __init__(self, *_args: Any, **_kwargs: Any) -> None:
            raise AssertionError("BANG_KE không được hỏi một tháng chung.")

    monkeypatch.setattr(main_window_module, "RepostSelectionDialog", Dialog)
    monkeypatch.setattr(
        main_window_module, "MonthSelectionDialog", UnexpectedMonthDialog
    )
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_context="workflow",
        _excel_operation="posting",
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow._excel_analysis_ready(owner, plan)

    assert tasks.cancel_calls == 1
    assert tasks.start_calls == [
        {"batch_id": 42, "repost_source_indices": {3}}
    ]


@pytest.mark.parametrize("selector_action", ["SELECT_SOURCE_ITEM", "SELECT_SHEET"])
def test_posting_selector_resolution_is_refined_before_apply(
    monkeypatch,
    selector_action: str,
) -> None:
    conflict = SimpleNamespace(conflict_id="selector-conflict")
    plan = SimpleNamespace(
        operation="posting",
        conflicts=[conflict],
        selected_sheet="T07 26",
        previously_posted_items=[],
    )

    class Tasks:
        def __init__(self) -> None:
            self.refine_calls: list[Any] = []
            self.apply_calls: list[Any] = []
            self.cancel_calls = 0

        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "posting"

        def refine_plan(
            self,
            value: Any,
            resolutions: Any,
            *,
            operation: str,
        ) -> None:
            self.refine_calls.append((value, resolutions, operation))

        def apply_plan(
            self,
            value: Any,
            resolutions: Any,
            *,
            operation: str,
        ) -> None:
            self.apply_calls.append((value, resolutions, operation))

        def cancel_waiting(self) -> None:
            self.cancel_calls += 1

    class Dialog:
        def __init__(
            self, _conflicts: Any, _parent: Any, **_kwargs: Any
        ) -> None:
            pass

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

        @staticmethod
        def resolution_map() -> dict[str, Any]:
            return {
                "selector-conflict": {
                    "conflict_id": "selector-conflict",
                    "action": selector_action,
                }
            }

    monkeypatch.setattr(main_window_module, "ConflictResolutionDialog", Dialog)
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_operation="posting",
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow._excel_analysis_ready(owner, plan)

    assert len(tasks.refine_calls) == 1
    assert tasks.refine_calls[0][1]["selector-conflict"]["action"] == selector_action
    assert tasks.apply_calls == []
    assert tasks.cancel_calls == 0


def test_bang_ke_non_selector_resolution_is_refined_before_apply(
    monkeypatch,
) -> None:
    conflict = SimpleNamespace(conflict_id="remaining-conflict")
    plan = SimpleNamespace(
        operation="posting",
        source_kind="BANG_KE",
        conflicts=[conflict],
        selected_sheet=None,
        previously_posted_items=[],
        confirmation_required=False,
    )

    class Tasks:
        def __init__(self) -> None:
            self.refine_calls: list[Any] = []
            self.apply_calls: list[Any] = []

        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "posting"

        def refine_plan(
            self,
            value: Any,
            resolutions: Any,
            *,
            operation: str,
        ) -> None:
            self.refine_calls.append((value, resolutions, operation))

        def apply_plan(
            self,
            value: Any,
            resolutions: Any,
            *,
            operation: str,
        ) -> None:
            self.apply_calls.append((value, resolutions, operation))

        @staticmethod
        def cancel_waiting() -> None:
            raise AssertionError("Plan phải được refine, không được hủy.")

    class Dialog:
        def __init__(
            self, _conflicts: Any, _parent: Any, **_kwargs: Any
        ) -> None:
            pass

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

        @staticmethod
        def resolution_map() -> dict[str, Any]:
            return {
                "remaining-conflict": {
                    "conflict_id": "remaining-conflict",
                    "action": "SKIP",
                }
            }

    monkeypatch.setattr(main_window_module, "ConflictResolutionDialog", Dialog)
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_operation="posting",
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow._excel_analysis_ready(owner, plan)

    assert tasks.refine_calls == [
        (
            plan,
            {
                "remaining-conflict": {
                    "conflict_id": "remaining-conflict",
                    "action": "SKIP",
                }
            },
            "posting",
        )
    ]
    assert tasks.apply_calls == []


def test_posting_reopens_dialog_with_saved_resolution(
    monkeypatch,
) -> None:
    conflict = SimpleNamespace(conflict_id="occupied")
    plan = SimpleNamespace(
        operation="posting",
        conflicts=[conflict],
        selected_sheet="T07 26",
        previously_posted_items=[],
    )

    class Tasks:
        def __init__(self) -> None:
            self.apply_calls: list[Any] = []

        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "posting"

        @staticmethod
        def saved_resolutions(_plan: Any, *, operation: str) -> dict[str, Any]:
            assert operation == "posting"
            return {
                "occupied": {
                    "conflict_id": "occupied",
                    "action": "KEEP_EXISTING",
                }
            }

        def apply_plan(
            self, value: Any, resolutions: Any, *, operation: str
        ) -> None:
            self.apply_calls.append((value, resolutions, operation))

        @staticmethod
        def cancel_waiting() -> None:
            raise AssertionError("Không được hủy phiên đã phục hồi.")

    class Dialog:
        initial_resolutions: dict[str, Any] = {}

        def __init__(self, _conflicts: Any, _parent: Any, **kwargs: Any) -> None:
            Dialog.initial_resolutions = dict(kwargs["initial_resolutions"])

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

        @staticmethod
        def resolution_map() -> dict[str, Any]:
            return {
                "occupied": {
                    "conflict_id": "occupied",
                    "action": "KEEP_EXISTING",
                }
            }

    monkeypatch.setattr(
        main_window_module, "ConflictResolutionDialog", Dialog
    )
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_operation="posting",
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow._excel_analysis_ready(owner, plan)

    assert Dialog.initial_resolutions["occupied"]["action"] == "KEEP_EXISTING"
    assert tasks.apply_calls == [
        (
            plan,
            {
                "occupied": {
                    "conflict_id": "occupied",
                    "action": "KEEP_EXISTING",
                }
            },
            "posting",
        )
    ]


def test_daily_sync_reopens_dialog_with_saved_resolution(monkeypatch) -> None:
    conflict = SimpleNamespace(conflict_id="daily-conflict")
    candidate = SimpleNamespace(
        month=7,
        target_sheet="T07 26",
        update_count=1,
        new_row_count=0,
        unchanged_count=0,
        target_only_count=0,
        invalid_count=0,
    )
    plan = SimpleNamespace(
        operation="sync",
        conflicts=[conflict],
        selected_month=7,
        selected_sheet="T07 26",
        month_candidates=[candidate],
    )

    class Tasks:
        def __init__(self) -> None:
            self.apply_calls: list[Any] = []

        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "sync"

        @staticmethod
        def saved_resolutions(_plan: Any, *, operation: str) -> dict[str, Any]:
            return {
                "daily-conflict": {
                    "conflict_id": "daily-conflict",
                    "action": "SKIP_INVALID",
                }
            }

        def apply_plan(
            self, value: Any, resolutions: Any, *, operation: str
        ) -> None:
            self.apply_calls.append((value, resolutions, operation))

        @staticmethod
        def cancel_waiting() -> None:
            raise AssertionError("Không được hủy phiên đã phục hồi.")

    class Dialog:
        initial_resolutions: dict[str, Any] = {}

        def __init__(self, _conflicts: Any, _parent: Any, **kwargs: Any) -> None:
            Dialog.initial_resolutions = dict(kwargs["initial_resolutions"])

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

        @staticmethod
        def resolution_map() -> dict[str, Any]:
            return {
                "daily-conflict": {
                    "conflict_id": "daily-conflict",
                    "action": "SKIP_INVALID",
                }
            }

    monkeypatch.setattr(
        main_window_module, "ConflictResolutionDialog", Dialog
    )
    class Confirmation:
        def __init__(self, *_args: Any, **_kwargs: Any) -> None:
            pass

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

    monkeypatch.setattr(main_window_module, "ExcelConfirmationDialog", Confirmation)
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_operation="sync",
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow._excel_analysis_ready(owner, plan)

    assert Dialog.initial_resolutions["daily-conflict"]["action"] == "SKIP_INVALID"
    assert tasks.apply_calls[0][1]["daily-conflict"]["action"] == "SKIP_INVALID"


def test_payment_sync_reopens_dialog_with_saved_resolution(
    monkeypatch,
) -> None:
    conflict = SimpleNamespace(conflict_id="payment-conflict")
    plan = SimpleNamespace(conflicts=[conflict])

    class Tasks:
        def __init__(self) -> None:
            self.refine_calls: list[Any] = []

        @staticmethod
        def saved_resolutions(_plan: Any, *, operation: str) -> dict[str, Any]:
            return {
                "payment-conflict": {
                    "conflict_id": "payment-conflict",
                    "action": "SKIP",
                }
            }

        def refine_plan(
            self, value: Any, resolutions: Any, *, operation: str
        ) -> None:
            self.refine_calls.append((value, resolutions, operation))

    class Dialog:
        initial_resolutions: dict[str, Any] = {}

        def __init__(self, _conflicts: Any, _parent: Any, **kwargs: Any) -> None:
            Dialog.initial_resolutions = dict(kwargs["initial_resolutions"])

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

        @staticmethod
        def resolution_map() -> dict[str, Any]:
            return {
                "payment-conflict": {
                    "conflict_id": "payment-conflict",
                    "action": "SKIP",
                }
            }

    monkeypatch.setattr(
        main_window_module, "ConflictResolutionDialog", Dialog
    )
    tasks = Tasks()
    owner = SimpleNamespace(_excel_tasks=tasks)

    MainWindow._handle_payment_sync_plan(owner, plan)

    assert Dialog.initial_resolutions["payment-conflict"]["action"] == "SKIP"
    assert tasks.refine_calls == [
        (
            plan,
            {
                "payment-conflict": {
                    "conflict_id": "payment-conflict",
                    "action": "SKIP",
                }
            },
            "payment_sync",
        )
    ]


def test_conflict_choices_are_prepared_before_final_confirmation(
    monkeypatch,
) -> None:
    conflict = SimpleNamespace(conflict_id="occupied")
    plan = SimpleNamespace(
        operation="posting",
        source_kind="BANG_KE",
        conflicts=[conflict],
        selected_sheet=None,
        previously_posted_items=[],
        confirmation_required=True,
        confirmation_done=False,
        original_source_count=1,
        reconciliation_source_count=0,
        items=[object()],
        target_sheets={"T01 26"},
    )
    expected = {
        "occupied": {
            "conflict_id": "occupied",
            "action": "KEEP_EXISTING",
        }
    }

    class Tasks:
        def __init__(self) -> None:
            self.saved: list[Any] = []
            self.cancelled = 0
            self.prepared: list[Any] = []

        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "posting"

        @staticmethod
        def saved_resolutions(_plan: Any, *, operation: str) -> dict[str, Any]:
            return {}

        def save_draft(
            self, value: Any, resolutions: Any, *, operation: str
        ) -> None:
            self.saved.append((value, resolutions, operation))

        def cancel_waiting(self) -> None:
            self.cancelled += 1

        def prepare_plan(
            self, value: Any, resolutions: Any, *, operation: str
        ) -> None:
            self.prepared.append((value, resolutions, operation))

    class Dialog:
        def __init__(
            self, _conflicts: Any, _parent: Any, **_kwargs: Any
        ) -> None:
            pass

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

        @staticmethod
        def resolution_map() -> dict[str, Any]:
            return expected

    monkeypatch.setattr(main_window_module, "ConflictResolutionDialog", Dialog)
    confirmation_calls: list[Any] = []
    monkeypatch.setattr(
        QMessageBox,
        "question",
        lambda *_args, **_kwargs: confirmation_calls.append(_args),
    )
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_operation="posting",
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow._excel_analysis_ready(owner, plan)

    assert tasks.saved == [(plan, expected, "posting")]
    assert tasks.prepared == [(plan, expected, "posting")]
    assert tasks.cancelled == 0
    assert confirmation_calls == []


def test_payment_sync_confirmation_discloses_new_sheet_template(
    monkeypatch,
) -> None:
    plan = SimpleNamespace(
        conflicts=[],
        new_rows=[],
        source_sheet="T06 26",
        target_sheet="T06 26",
        target_sheet_created=True,
        template_sheet="T05 26",
        update_count=0,
        unchanged_count=0,
        new_count=0,
        conflict_count=0,
        normalization_sheet_count=0,
        targets={
            "HP": SimpleNamespace(
                sheet_name="T06 26 HP", sheet_to_create=True,
                template_sheet="T05 26 HP", new_count=1,
                update_count=2, unchanged_count=3, conflict_count=0,
            ),
            "NAM": SimpleNamespace(
                sheet_name="T06 26 NAM", sheet_to_create=True,
                template_sheet="T05 26 NAM", new_count=4,
                update_count=5, unchanged_count=6, conflict_count=1,
            ),
        },
    )

    class Tasks:
        def __init__(self) -> None:
            self.apply_calls: list[Any] = []

        def apply_plan(
            self,
            value: Any,
            resolutions: Any,
            *,
            operation: str,
        ) -> None:
            self.apply_calls.append((value, resolutions, operation))

        @staticmethod
        def cancel_waiting() -> None:
            raise AssertionError("Không được hủy khi người dùng xác nhận.")

    class Confirmation:
        summary: Any = None

        def __init__(self, summary: Any, **_kwargs: Any) -> None:
            Confirmation.summary = summary

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

    monkeypatch.setattr(main_window_module, "ExcelConfirmationDialog", Confirmation)
    tasks = Tasks()
    owner = SimpleNamespace(_excel_tasks=tasks)

    MainWindow._handle_payment_sync_plan(owner, plan)

    assert len(tasks.apply_calls) == 1
    assert Confirmation.summary.title == "Xác nhận đồng bộ BK → Thanh toán"
    assert Confirmation.summary.tables[0].rows == (
        ("T06 26", "T06 26 HP", "Tạo từ T05 26 HP", 2, 1, 3, 0),
        ("T06 26", "T06 26 NAM", "Tạo từ T05 26 NAM", 5, 4, 6, 0),
    )
    assert "T05 26 HP" in Confirmation.summary.notices[0].text
    assert "T05 26 NAM" in Confirmation.summary.notices[1].text


@pytest.mark.parametrize(
    ("operation", "result_fields", "expected_title", "expected_metric"),
    (
        (
            "sync",
            {
                "message": "Đồng bộ xong.",
                "sheet_name": "T07 26",
                "updated_rows": 2,
                "inserted_rows": 1,
                "target_only_rows": 3,
                "invalid_rows": 0,
            },
            "Đồng bộ Hàng ngày → BK hoàn tất",
            2,
        ),
        (
            "posting",
            {
                "message": "Nhập khoản chi xong.",
                "sheet_name": "T07 26",
                "posted_source_items": 10,
                "already_existing_items": 2,
                "skipped_source_items": 1,
            },
            "Nhập khoản chi vào BK hoàn tất",
            10,
        ),
    ),
)
def test_bk_completion_can_open_the_written_bk_file(
    monkeypatch,
    tmp_path: Path,
    operation: str,
    result_fields: dict[str, Any],
    expected_title: str,
    expected_metric: int,
) -> None:
    bk_path = tmp_path / "BK 2026.xlsx"
    bk_path.touch()
    opened: list[Path] = []

    class CompletionDialog:
        OPEN_TARGET = "open_target"
        VIEW_DETAILS = "view_details"
        summary: Any = None

        def __init__(self, summary: Any, **_kwargs: Any) -> None:
            CompletionDialog.summary = summary
            self.selected_action = self.OPEN_TARGET

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

    class Tasks:
        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return operation

    monkeypatch.setattr(main_window_module, "ExcelCompletionDialog", CompletionDialog)
    owner = SimpleNamespace(
        _excel_context="workflow",
        _excel_operation=operation,
        _excel_tasks=Tasks(),
        workflow_page=SimpleNamespace(
            set_excel_result=lambda *_args: None
        ),
        _load_excel_history=lambda: None,
        _open_bk_workbook=lambda path: opened.append(Path(path)),
    )
    result = SimpleNamespace(
        operation=operation,
        target_path=bk_path,
        **result_fields,
    )

    MainWindow._excel_completed(owner, result)

    assert CompletionDialog.summary.title == expected_title
    assert CompletionDialog.summary.metrics[0].value == expected_metric
    assert opened == [bk_path]


def test_payment_completion_displays_carrier_invoice_summary(
    monkeypatch,
    tmp_path: Path,
) -> None:
    payment_path = tmp_path / "Thanh toan.xlsm"
    payment_path.touch()
    opened: list[Path] = []

    class CompletionDialog:
        OPEN_TARGET = "open_target"
        VIEW_DETAILS = "view_details"
        summary: Any = None

        def __init__(self, summary: Any, **_kwargs: Any) -> None:
            CompletionDialog.summary = summary
            self.selected_action = self.OPEN_TARGET

        @staticmethod
        def exec() -> QDialog.DialogCode:
            return QDialog.DialogCode.Accepted

    class Tasks:
        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "payment_sync"

    monkeypatch.setattr(main_window_module, "ExcelCompletionDialog", CompletionDialog)
    owner = SimpleNamespace(
        _excel_context="workflow",
        _excel_operation="payment_sync",
        _excel_tasks=Tasks(),
        workflow_page=SimpleNamespace(set_excel_result=lambda *_args: None),
        _load_excel_history=lambda: None,
        _open_workbook_path=lambda path, **_kwargs: opened.append(Path(path)),
    )
    result = SimpleNamespace(
        operation="payment_sync",
        message="Hoàn tất.",
        target_path=payment_path,
        source_sheet_name="T07 26",
        sheet_name="T07 26 HP, T07 26 NAM",
        target_results={},
        carrier_summary=SimpleNamespace(
            period="T07/2026",
            carrier_count=2,
            invoice_count=3,
            selected_period_total=1_700_000,
            missing_invoice_items=1,
            unmapped_carrier_items=0,
            cross_carrier_invoices=1,
            suspected_duplicate_items=1,
        ),
    )

    MainWindow._excel_completed(owner, result)

    details = CompletionDialog.summary.details[0]
    assert details.title == "Tổng hợp hóa đơn và bên vận tải"
    assert "Số hóa đơn: 3" in details.lines
    assert "Tổng tiền kỳ: 1.700.000" in details.lines
    assert "Hóa đơn thuộc nhiều bên: 1" in details.lines
    assert details.expanded is True
    assert opened == [payment_path]


def test_open_bk_workbook_uses_the_default_desktop_application(
    monkeypatch,
    tmp_path: Path,
) -> None:
    bk_path = tmp_path / "BK 2026.xlsx"
    bk_path.touch()
    opened_urls: list[Any] = []
    monkeypatch.setattr(
        main_window_module.QDesktopServices,
        "openUrl",
        lambda url: opened_urls.append(url) or True,
    )
    owner = SimpleNamespace(_settings=None)

    MainWindow._open_bk_workbook(owner, bk_path)

    assert len(opened_urls) == 1
    assert Path(opened_urls[0].toLocalFile()) == bk_path


def test_settings_round_trip_optional_excel_paths_and_reject_xls(
    qtbot,
    tmp_path: Path,
) -> None:
    bat = tmp_path / "assistant.bat"
    bat.write_text("@echo off", encoding="utf-8")
    daily = tmp_path / "Hang ngay 2026.xlsx"
    bk = tmp_path / "BK 2026.xlsm"
    page = SettingsPage(
        {
            "assistant_bat_path": str(bat),
            "output_dir": str(tmp_path),
            "daily_workbook_path": str(daily),
            "bk_workbook_path": str(bk),
        }
    )
    qtbot.addWidget(page)

    assert page.save_button.isEnabled()
    assert page.settings_data()["daily_workbook_path"] == str(daily)
    assert page.settings_data()["bk_workbook_path"] == str(bk)

    page.daily_workbook_edit.setText(str(tmp_path / "legacy.xls"))

    assert not page.save_button.isEnabled()
    assert ".xlsx hoặc .xlsm" in page.validation_label.text()


class _Service:
    def __init__(self, plan: Any) -> None:
        self.plan = plan
        self.apply_calls: list[tuple[Any, Any]] = []

    def analyze(self, *, progress_callback) -> Any:
        progress_callback("Đang phân tích")
        return self.plan

    def apply(self, plan: Any, resolutions: Any, *, progress_callback) -> Any:
        self.apply_calls.append((plan, resolutions))
        progress_callback("Đang ghi")
        return {"message": "Hoàn tất", "operation": "sync"}


class _DraftService:
    def __init__(self) -> None:
        self.saved: list[Any] = []
        self.completed: list[str] = []
        self.failed: list[tuple[str, str]] = []
        self.cancelled: list[str] = []

    @staticmethod
    def restore_with_info(_plan: Any, _operation: str) -> Any:
        return SimpleNamespace(
            source_file_key="posting:batch:42",
            resolutions={"c1": {"conflict_id": "c1", "action": "SKIP"}},
            found=True,
            status="FAILED",
            updated_at="2026-08-14T13:57:43+07:00",
            current_conflict_count=1,
            saved_choice_count=1,
            restored_count=1,
            target_compatible=True,
            last_error="temporary error",
        )

    def save(self, plan: Any, operation: str, resolutions: Any) -> str:
        self.saved.append((plan, operation, resolutions))
        return "posting:batch:42"

    def mark_completed(self, key: str) -> None:
        self.completed.append(key)

    def mark_failed(self, key: str, error: object) -> None:
        self.failed.append((key, str(error)))

    def mark_cancelled(self, key: str) -> None:
        self.cancelled.append(key)


def test_excel_controller_auto_applies_plan_without_conflicts(qtbot) -> None:
    service = _Service({"operation": "sync", "has_changes": True, "conflicts": []})
    controller = ExcelTaskController(daily_sync_service=service)
    completed: list[Any] = []
    progress: list[tuple[str, str]] = []
    controller.completed.connect(completed.append)
    controller.progress.connect(lambda operation, message: progress.append((operation, message)))

    try:
        controller.start_sync()
        qtbot.waitUntil(lambda: bool(completed), timeout=2000)

        assert service.apply_calls
        assert service.apply_calls[0][1] == {}
        assert ("sync", "Đang phân tích") in progress
        assert ("sync", "Đang ghi") in progress
        assert not controller.is_busy
    finally:
        controller.shutdown()


def test_excel_controller_waits_for_aggregate_resolution(qtbot) -> None:
    plan = {
        "operation": "posting",
        "conflicts": [{"conflict_id": "c1", "type": "TARGET_CELL_OCCUPIED"}],
    }
    service = _Service(plan)
    controller = ExcelTaskController(expense_posting_service=service)
    ready: list[Any] = []
    completed: list[Any] = []
    controller.analysis_ready.connect(ready.append)
    controller.completed.connect(completed.append)

    try:
        controller.start_posting()
        qtbot.waitUntil(lambda: bool(ready), timeout=2000)

        assert controller.phase == "waiting_user"
        with pytest.raises(RuntimeError):
            controller.start_sync()

        controller.apply_plan(
            ready[0],
            {"c1": {"conflict_id": "c1", "action": "KEEP_EXISTING"}},
        )
        qtbot.waitUntil(lambda: bool(completed), timeout=2000)

        assert service.apply_calls[0][1]["c1"]["action"] == "KEEP_EXISTING"
    finally:
        controller.shutdown()


def test_excel_controller_refines_selector_before_apply(qtbot) -> None:
    initial = {
        "operation": "posting",
        "conflicts": [{"conflict_id": "row", "type": "BL_ONLY_NO_CONTAINER"}],
    }

    class RefiningService(_Service):
        def __init__(self) -> None:
            super().__init__(initial)
            self.refine_calls: list[Any] = []

        def refine(
            self, plan: Any, resolutions: Any, *, progress_callback
        ) -> Any:
            self.refine_calls.append((plan, resolutions))
            progress_callback("Đang kiểm tra lại ô")
            return {
                "operation": "posting",
                "has_changes": True,
                "conflicts": [],
            }

        def apply(
            self, plan: Any, resolutions: Any, *, progress_callback
        ) -> Any:
            self.apply_calls.append((plan, resolutions))
            return {"message": "Hoàn tất", "operation": "posting"}

    service = RefiningService()
    controller = ExcelTaskController(expense_posting_service=service)
    ready: list[Any] = []
    completed: list[Any] = []
    controller.analysis_ready.connect(ready.append)
    controller.completed.connect(completed.append)

    try:
        controller.start_posting()
        qtbot.waitUntil(lambda: bool(ready), timeout=2000)
        controller.refine_plan(
            ready[0],
            {
                "row": {
                    "conflict_id": "row",
                    "action": "SELECT_ROW",
                    "selected_row": 12,
                }
            },
        )
        qtbot.waitUntil(lambda: bool(completed), timeout=2000)

        assert service.refine_calls
        assert service.apply_calls
        assert service.apply_calls[0][1] == {}
    finally:
        controller.shutdown()


def test_excel_controller_notifies_service_when_user_cancels(qtbot) -> None:
    plan = {
        "operation": "posting",
        "conflicts": [{"conflict_id": "c1", "type": "CONTAINER_NOT_FOUND"}],
    }

    class CancellableService(_Service):
        def __init__(self) -> None:
            super().__init__(plan)
            self.cancelled: list[Any] = []

        def cancel(self, value: Any) -> None:
            self.cancelled.append(value)

    service = CancellableService()
    controller = ExcelTaskController(expense_posting_service=service)
    ready: list[Any] = []
    controller.analysis_ready.connect(ready.append)

    try:
        controller.start_posting()
        qtbot.waitUntil(lambda: bool(ready), timeout=2000)

        assert controller.cancel_waiting()
        assert service.cancelled == [plan]
        assert not controller.is_busy
    finally:
        controller.shutdown()


def test_controller_keeps_successful_choices_and_restore_metadata(qtbot) -> None:
    plan = {
        "operation": "posting",
        "conflicts": [{"conflict_id": "c1", "type": "CONTAINER_NOT_FOUND"}],
    }
    service = _Service(plan)
    drafts = _DraftService()
    controller = ExcelTaskController(
        expense_posting_service=service,
        draft_service=drafts,
    )
    ready: list[Any] = []
    completed: list[Any] = []
    controller.analysis_ready.connect(ready.append)
    controller.completed.connect(completed.append)

    try:
        controller.start_posting()
        qtbot.waitUntil(lambda: bool(ready), timeout=2000)
        restored = controller.saved_resolutions(ready[0], operation="posting")

        assert restored["c1"]["action"] == "SKIP"
        assert controller.saved_resolution_info["status"] == "FAILED"

        controller.apply_plan(ready[0], restored, operation="posting")
        qtbot.waitUntil(lambda: bool(completed), timeout=2000)

        assert drafts.saved
        assert drafts.completed == ["posting:batch:42"]
    finally:
        controller.shutdown()


def test_controller_does_not_overwrite_history_when_dialog_is_only_closed(
    qtbot,
) -> None:
    plan = {
        "operation": "posting",
        "conflicts": [{"conflict_id": "c1", "type": "CONTAINER_NOT_FOUND"}],
    }
    service = _Service(plan)
    drafts = _DraftService()
    controller = ExcelTaskController(
        expense_posting_service=service,
        draft_service=drafts,
    )
    ready: list[Any] = []
    controller.analysis_ready.connect(ready.append)

    try:
        controller.start_posting()
        qtbot.waitUntil(lambda: bool(ready), timeout=2000)
        controller.saved_resolutions(ready[0], operation="posting")

        assert controller.cancel_waiting()
        assert drafts.cancelled == []
    finally:
        controller.shutdown()


def test_controller_marks_choices_cancelled_after_they_were_saved(qtbot) -> None:
    plan = {
        "operation": "posting",
        "conflicts": [{"conflict_id": "c1", "type": "CONTAINER_NOT_FOUND"}],
    }
    service = _Service(plan)
    drafts = _DraftService()
    controller = ExcelTaskController(
        expense_posting_service=service,
        draft_service=drafts,
    )
    ready: list[Any] = []
    controller.analysis_ready.connect(ready.append)

    try:
        controller.start_posting()
        qtbot.waitUntil(lambda: bool(ready), timeout=2000)
        controller.save_draft(
            ready[0],
            {"c1": {"conflict_id": "c1", "action": "SKIP"}},
            operation="posting",
        )

        assert controller.cancel_waiting()
        assert drafts.cancelled == ["posting:batch:42"]
    finally:
        controller.shutdown()


def test_negative_adjustment_requires_action_before_dialog_accepts(qtbot) -> None:
    dialog = ConflictResolutionDialog(
        [
            {
                "conflict_id": "negative",
                "type": "NEGATIVE_ADJUSTMENT",
                "container": "ABCD1234567",
                "fee": "NH",
                "amount": -100,
                "allowed_actions": ["ADD", "SKIP", "CANCEL_ALL"],
            }
        ]
    )
    qtbot.addWidget(dialog)

    dialog._validate_and_accept()

    assert dialog.result() != QDialog.DialogCode.Accepted
    assert "chưa chọn cách xử lý" in dialog.validation_label.text()

    dialog._action_combos["negative"].setCurrentIndex(1)
    dialog._validate_and_accept()

    assert dialog.result() == QDialog.DialogCode.Accepted
    assert dialog.resolution_map()["negative"]["action"] == "ADD"


def test_month_and_conflict_dialogs_collect_generic_mapping(qtbot) -> None:
    months = MonthSelectionDialog(
        [
            {"sheet_name": "T06 26", "month": 6, "year": 2026, "match_count": 2},
            {
                "sheet_name": "T07 26",
                "month": 7,
                "year": 2026,
                "match_count": 4,
                "is_recent": True,
            },
        ]
    )
    qtbot.addWidget(months)
    months.table.selectRow(1)

    assert months.selected_sheet_name == "T07 26"
    assert months.selection()["selected_sheet_name"] == "T07 26"

    conflicts = ConflictResolutionDialog(
        [
            {
                "conflict_id": "occupied",
                "type": "TARGET_CELL_OCCUPIED",
                "container": "DRYU3026167",
                "fee": "VTN",
                "amount": 1_000_000,
                "target_cell": "Q12",
                "current_value": 800_000,
            },
            {
                "conflict_id": "unknown",
                "type": "UNKNOWN_FEE_CODE",
                "fee": "CXD",
                "amount": 500_000,
            },
        ]
    )
    qtbot.addWidget(conflicts)
    occupied_combo = conflicts._action_combos["occupied"]
    occupied_combo.setCurrentIndex(occupied_combo.findData("OVERWRITE"))
    fee_action = conflicts._action_combos["unknown"]
    fee_action.setCurrentIndex(fee_action.findData("SELECT_FEE"))
    conflicts._selected_fees["unknown"].setCurrentIndex(
        conflicts._selected_fees["unknown"].findData("SC")
    )

    result = conflicts.resolution_map()

    assert result["occupied"]["action"] == "OVERWRITE"
    assert result["unknown"]["action"] == "SELECT_FEE"
    assert result["unknown"]["selected_fee"] == "SC"


def test_conflict_dialog_prefills_saved_resolutions(qtbot) -> None:
    dialog = ConflictResolutionDialog(
        [
            {
                "conflict_id": "occupied",
                "type": "TARGET_CELL_OCCUPIED",
                "target_cell": "Q12",
                "current_value": 800_000,
            },
            {
                "conflict_id": "unknown",
                "type": "UNKNOWN_FEE_CODE",
                "fee": "CXD",
                "amount": 500_000,
            },
        ],
        initial_resolutions={
            "occupied": {
                "conflict_id": "occupied",
                "action": "OVERWRITE",
            },
            "unknown": {
                "conflict_id": "unknown",
                "action": "SELECT_FEE",
                "selected_fee": "SC",
            },
        },
        restore_info={
            "found": True,
            "status": "FAILED",
            "updated_at": "2026-08-14T13:57:43+07:00",
            "target_compatible": True,
        },
    )
    qtbot.addWidget(dialog)

    assert dialog._action_combos["occupied"].currentData() == "OVERWRITE"
    assert dialog._action_combos["unknown"].currentData() == "SELECT_FEE"
    assert dialog._selected_fees["unknown"].currentData() == "SC"
    assert "Đã khôi phục 2/2 lựa chọn" in dialog.restore_label.text()
    assert "Thất bại" in dialog.restore_label.text()


def test_conflict_dialog_restores_selected_row_and_source_sheet(qtbot) -> None:
    dialog = ConflictResolutionDialog(
        [
            {
                "conflict_id": "selected-row",
                "type": "MULTIPLE_CONTAINER_MATCH",
                "container": "ABCD1234567",
                "fee": "VSDL",
                "amount": 270_000,
                "allowed_actions": ["SELECT_ROW", "SKIP", "CANCEL_ALL"],
                "row_candidates": [
                    {
                        "source_sheet": "T01 26",
                        "row": 55,
                        "container": "ABCD1234567",
                    }
                ],
            }
        ],
        initial_resolutions={
            "selected-row": {
                "conflict_id": "selected-row",
                "action": "SELECT_ROW",
                "selected_row": 55,
                "row": 55,
                "selected_source_sheet": "T01 26",
            }
        },
    )
    qtbot.addWidget(dialog)

    assert dialog._action_combos["selected-row"].currentData() == "SELECT_ROW"
    assert dialog._selected_rows["selected-row"] == 55
    assert dialog._selected_source_sheets["selected-row"] == "T01 26"
    assert dialog._selector_buttons["selected-row"].text() == "T01 26 – dòng 55"
    assert dialog.resolution_map()["selected-row"] == {
        "conflict_id": "selected-row",
        "action": "SELECT_ROW",
        "selected_row": 55,
        "row": 55,
        "selected_source_sheet": "T01 26",
    }

    dialog._validate_and_accept()
    assert dialog.result() == QDialog.DialogCode.Accepted


def test_posting_month_dialog_has_no_default_or_recommendation(qtbot) -> None:
    dialog = MonthSelectionDialog(
        [
            {
                "sheet_name": "T06 26",
                "month": 6,
                "year": 2026,
                "match_count": 2,
                "is_recent": True,
            },
            {
                "sheet_name": "T07 26",
                "month": 7,
                "year": 2026,
                "match_count": 4,
            },
        ],
        title="Chọn sheet nhận khoản chi",
        preselect_first=False,
        show_recommendations=False,
    )
    qtbot.addWidget(dialog)

    assert dialog.table.currentRow() == -1
    assert dialog.selected_sheet_name is None
    assert dialog.table.isColumnHidden(3)
    assert dialog.table.isColumnHidden(4)
    assert not dialog.buttons.button(
        QDialogButtonBox.StandardButton.Ok
    ).isEnabled()

    dialog.table.selectRow(1)

    assert dialog.selected_sheet_name == "T07 26"
    assert dialog.buttons.button(
        QDialogButtonBox.StandardButton.Ok
    ).isEnabled()


def test_repost_dialog_defaults_to_unposted_and_selects_individual_rows(
    qtbot,
) -> None:
    dialog = RepostSelectionDialog(
        [
            {
                "source_item_index": 3,
                "container": "DRYU3026167",
                "fee_selected": "VTN",
                "amount": 1_000_000,
                "sheet_name": "T07 26",
                "target_row": 12,
                "target_cell": "Q12",
                "created_at": "2026-07-29T09:15:00",
            }
        ]
    )
    qtbot.addWidget(dialog)

    assert dialog.unposted_only.isChecked()
    assert not dialog.table.isEnabled()
    assert not dialog.select_all.isEnabled()
    assert dialog.selected_source_indices == []

    dialog.choose_reposts.setChecked(True)
    dialog.table.item(0, 0).setCheckState(Qt.CheckState.Checked)

    assert dialog.table.isEnabled()
    assert dialog.select_all.isEnabled()
    assert dialog.selected_source_indices == [3]


def test_repost_dialog_select_all_tracks_all_row_checkboxes(qtbot) -> None:
    dialog = RepostSelectionDialog(
        [
            {"source_item_index": 1},
            {"source_item_index": 3},
            {"source_item_index": 5},
        ]
    )
    qtbot.addWidget(dialog)
    dialog.choose_reposts.setChecked(True)

    dialog.select_all.setChecked(True)

    assert dialog.selected_source_indices == [1, 3, 5]
    assert all(
        dialog.table.item(row, 0).checkState() is Qt.CheckState.Checked
        for row in range(dialog.table.rowCount())
    )

    dialog.table.item(1, 0).setCheckState(Qt.CheckState.Unchecked)

    assert dialog.select_all.checkState() is Qt.CheckState.PartiallyChecked
    assert dialog.selected_source_indices == [1, 5]

    dialog.select_all.setChecked(False)

    assert dialog.selected_source_indices == []
    assert all(
        dialog.table.item(row, 0).checkState() is Qt.CheckState.Unchecked
        for row in range(dialog.table.rowCount())
    )


def test_conflict_actions_are_short_vietnamese_labels(qtbot) -> None:
    dialog = ConflictResolutionDialog(
        [
            {
                "conflict_id": "duplicate",
                "type": "MULTIPLE_CONTAINER_MATCH",
                "allowed_actions": ["SKIP", "SELECT_ROW"],
                "row_candidates": [{"row_number": 12}, {"row_number": 13}],
            },
            {
                "conflict_id": "occupied",
                "type": "TARGET_CELL_OCCUPIED",
                "allowed_actions": [
                    "KEEP_EXISTING",
                    "OVERWRITE",
                    "SKIP",
                ],
            },
            {
                "conflict_id": "repeated",
                "type": "REPEATED_SOURCE_CONTAINER",
                "allowed_actions": ["SKIP", "SELECT_ROW"],
                "row_candidates": [{"row_number": 12, "sqt": 700}],
            },
        ]
    )
    qtbot.addWidget(dialog)

    duplicate_labels = [
        dialog._action_combos["duplicate"].itemText(index)
        for index in range(dialog._action_combos["duplicate"].count())
    ]
    occupied_labels = [
        dialog._action_combos["occupied"].itemText(index)
        for index in range(dialog._action_combos["occupied"].count())
    ]

    assert duplicate_labels == ["", "Chọn dòng", "Bỏ qua"]
    assert occupied_labels == ["", "Giữ nguyên", "Ghi đè"]
    assert "repeated" in dialog._selector_buttons
    assert not any(
        "_" in label
        for label in duplicate_labels + occupied_labels
    )


def test_conflict_bulk_action_applies_to_all_compatible_rows_and_allows_override(
    qtbot,
) -> None:
    dialog = ConflictResolutionDialog(
        [
            {
                "conflict_id": "row-a",
                "type": "MULTIPLE_CONTAINER_MATCH",
                "allowed_actions": ["SELECT_ROW", "SKIP"],
                "row_candidates": [{"row_number": 12}],
            },
            {
                "conflict_id": "row-b",
                "type": "CONTAINER_NOT_FOUND",
                "allowed_actions": ["SELECT_ROW", "SKIP"],
                "row_candidates": [{"row_number": 13}],
            },
            {
                "conflict_id": "occupied",
                "type": "TARGET_CELL_OCCUPIED",
                "allowed_actions": ["KEEP_EXISTING", "OVERWRITE"],
            },
        ]
    )
    qtbot.addWidget(dialog)

    skip_index = next(
        index
        for index in range(dialog.bulk_action_combo.count())
        if "SKIP" in dialog.bulk_action_combo.itemData(index)
    )
    dialog.bulk_action_combo.setCurrentIndex(skip_index)

    assert dialog.bulk_apply_button.isEnabled()
    assert dialog.bulk_action_combo.currentText() == "Bỏ qua (2 dòng)"

    dialog.bulk_apply_button.click()

    assert dialog._action_combos["row-a"].currentData() == "SKIP"
    assert dialog._action_combos["row-b"].currentData() == "SKIP"
    assert dialog._action_combos["occupied"].currentData() == ""
    assert "2/3 dòng" in dialog.bulk_result_label.text()

    row_a = dialog._action_combos["row-a"]
    row_a.setCurrentIndex(row_a.findData("SELECT_ROW"))

    assert row_a.currentData() == "SELECT_ROW"
    assert dialog._action_combos["row-b"].currentData() == "SKIP"


def test_conflict_bulk_action_applies_only_to_selected_rows_after_sorting(
    qtbot,
) -> None:
    dialog = ConflictResolutionDialog(
        [
            {
                "conflict_id": "row-b",
                "type": "TARGET_CELL_OCCUPIED",
                "container": "VSCU0000002",
                "allowed_actions": ["KEEP_EXISTING", "OVERWRITE"],
            },
            {
                "conflict_id": "row-a",
                "type": "TARGET_CELL_OCCUPIED",
                "container": "VSCU0000001",
                "allowed_actions": ["KEEP_EXISTING", "OVERWRITE"],
            },
            {
                "conflict_id": "row-c",
                "type": "TARGET_CELL_OCCUPIED",
                "container": "VSCU0000003",
                "allowed_actions": ["KEEP_EXISTING", "OVERWRITE"],
            },
        ]
    )
    qtbot.addWidget(dialog)
    dialog.table.sortItems(0, Qt.SortOrder.AscendingOrder)

    selection = dialog.table.selectionModel()
    flags = (
        QItemSelectionModel.SelectionFlag.Select
        | QItemSelectionModel.SelectionFlag.Rows
    )
    selection.select(dialog.table.model().index(0, 0), flags)
    selection.select(dialog.table.model().index(2, 0), flags)

    overwrite_index = next(
        index
        for index in range(dialog.bulk_action_combo.count())
        if "OVERWRITE" in dialog.bulk_action_combo.itemData(index)
    )
    dialog.bulk_action_combo.setCurrentIndex(overwrite_index)

    assert dialog.bulk_apply_selected_button.isEnabled()
    assert dialog.bulk_apply_selected_button.text().endswith("(2)")
    assert dialog._selected_conflict_ids() == {"row-a", "row-c"}

    dialog.bulk_apply_selected_button.click()

    assert dialog._action_combos["row-a"].currentData() == "OVERWRITE"
    assert dialog._action_combos["row-c"].currentData() == "OVERWRITE"
    assert dialog._action_combos["row-b"].currentData() == ""
    assert "2 dòng đã chọn" in dialog.bulk_result_label.text()


def test_conflict_selected_bulk_applies_select_row_but_still_requires_bk_row(
    qtbot,
) -> None:
    dialog = ConflictResolutionDialog(
        [
            {
                "conflict_id": "row-a",
                "type": "MULTIPLE_CONTAINER_MATCH",
                "allowed_actions": ["SELECT_ROW", "SKIP"],
                "row_candidates": [{"row_number": 12}],
            },
            {
                "conflict_id": "row-b",
                "type": "MULTIPLE_CONTAINER_MATCH",
                "allowed_actions": ["SELECT_ROW", "SKIP"],
                "row_candidates": [{"row_number": 13}],
            },
        ]
    )
    qtbot.addWidget(dialog)
    row_b = dialog._action_combos["row-b"]
    row_b.setCurrentIndex(row_b.findData("SKIP"))
    dialog.table.selectRow(0)

    select_row_index = next(
        index
        for index in range(dialog.bulk_action_combo.count())
        if "SELECT_ROW" in dialog.bulk_action_combo.itemData(index)
    )
    dialog.bulk_action_combo.setCurrentIndex(select_row_index)

    assert dialog.bulk_apply_selected_button.isEnabled()
    assert dialog.bulk_apply_button.isEnabled()

    dialog.bulk_apply_selected_button.click()

    assert dialog._action_combos["row-a"].currentData() == "SELECT_ROW"
    assert dialog._selector_buttons["row-a"].isEnabled()
    assert dialog._action_combos["row-b"].currentData() == "SKIP"
    assert not dialog._selector_buttons["row-b"].isEnabled()

    dialog._validate_and_accept()

    assert dialog.result() != QDialog.DialogCode.Accepted
    assert "chưa chọn dòng BK" in dialog.validation_label.text()


def test_expense_conflict_dialog_shows_source_vessel_voyage_only_for_posting(
    qtbot,
) -> None:
    conflict = {
        "conflict_id": "occupied",
        "type": "TARGET_CELL_OCCUPIED",
        "vessel_voyage": "NEW VISION 2610S",
        "allowed_actions": ["KEEP_EXISTING", "OVERWRITE"],
    }
    posting = ConflictResolutionDialog([conflict], operation="posting")
    daily = ConflictResolutionDialog([conflict], operation="sync")
    qtbot.addWidget(posting)
    qtbot.addWidget(daily)
    column = posting.COLUMNS.index("Tàu / chuyến nguồn")

    assert not posting.table.isColumnHidden(column)
    assert posting.table.item(0, column).text() == "NEW VISION 2610S"
    assert daily.table.isColumnHidden(column)


def test_skipping_row_selection_disables_and_clears_target_value_choice(
    qtbot,
) -> None:
    dialog = ConflictResolutionDialog(
        [
            {
                "conflict_id": "row",
                "item_index": 4,
                "type": "MULTIPLE_CONTAINER_MATCH",
                "allowed_actions": ["SELECT_ROW", "SKIP"],
                "row_candidates": [
                    {"source_sheet": "T07 26", "row_number": 18}
                ],
            },
            {
                "conflict_id": "cell",
                "item_index": 4,
                "type": "TARGET_CELL_OCCUPIED",
                "allowed_actions": ["KEEP_EXISTING", "OVERWRITE", "SKIP"],
            },
        ],
        initial_resolutions={
            "row": {
                "action": "SELECT_ROW",
                "selected_source_sheet": "T07 26",
                "selected_row": 18,
            },
            "cell": {"action": "OVERWRITE"},
        },
    )
    qtbot.addWidget(dialog)

    row_combo = dialog._action_combos["row"]
    cell_combo = dialog._action_combos["cell"]
    assert row_combo.currentData() == "SELECT_ROW"
    assert cell_combo.currentData() == "OVERWRITE"
    assert cell_combo.isEnabled()

    row_combo.setCurrentIndex(row_combo.findData("SKIP"))

    assert not cell_combo.isEnabled()
    assert cell_combo.currentData() == ""
    assert dialog._selected_rows["row"] is None
    assert dialog._selector_buttons["row"].text() == "Chọn dòng…"
    dialog._validate_and_accept()
    assert dialog.result() == QDialog.DialogCode.Accepted


def test_target_value_choice_starts_blank_and_is_required(qtbot) -> None:
    dialog = ConflictResolutionDialog(
        [
            {
                "conflict_id": "occupied",
                "type": "TARGET_CELL_OCCUPIED",
                "allowed_actions": ["KEEP_EXISTING", "OVERWRITE", "SKIP"],
            }
        ]
    )
    qtbot.addWidget(dialog)

    combo = dialog._action_combos["occupied"]
    assert [combo.itemText(index) for index in range(combo.count())] == [
        "",
        "Giữ nguyên",
        "Ghi đè",
    ]
    assert combo.currentData() == ""

    dialog._validate_and_accept()
    assert dialog.result() != QDialog.DialogCode.Accepted
    assert "chưa chọn cách xử lý" in dialog.validation_label.text()


def test_invoice_conflict_dialog_selects_one_invoice_and_has_only_two_value_actions(
    qtbot,
) -> None:
    dialog = ConflictResolutionDialog(
        [
            {
                "conflict_id": "multiple-invoices",
                "type": "MULTIPLE_SOURCE_INVOICES",
                "allowed_actions": ["SELECT_INVOICE"],
                "details": {"invoice_candidates": ["INV-A", "INV-B"]},
            },
            {
                "conflict_id": "invoice-value",
                "type": "INVOICE_VALUE_CONFLICT",
                "allowed_actions": ["KEEP_EXISTING", "OVERWRITE"],
                "current_value": "INV-OLD",
                "details": {"invoice_candidates": ["INV-NEW"]},
            },
        ]
    )
    qtbot.addWidget(dialog)

    assert "Số HĐ từ JSON" in dialog.COLUMNS
    value_labels = [
        dialog._action_combos["invoice-value"].itemText(index)
        for index in range(dialog._action_combos["invoice-value"].count())
    ]
    assert value_labels == ["Giữ HĐ hiện tại", "Ghi đè bằng HĐ mới"]
    invoice_combo = dialog._selected_invoices["multiple-invoices"]
    assert invoice_combo.currentData() is None

    dialog._validate_and_accept()
    assert dialog.result() == QDialog.DialogCode.Rejected
    assert "chưa chọn Số HĐ" in dialog.validation_label.text()

    invoice_combo.setCurrentIndex(invoice_combo.findData("INV-B"))
    result = dialog.resolution_map()
    assert result["multiple-invoices"] == {
        "conflict_id": "multiple-invoices",
        "action": "SELECT_INVOICE",
        "selected_invoice": "INV-B",
    }


def test_manual_row_picker_returns_source_sheet_and_workbook_row(qtbot) -> None:
    dialog = ManualRowPickerDialog(
        [
                {
                    "source_sheet": "T06 26",
                "row_number": 12,
                "sqt": 700,
                "container": "DRYU3026167",
                "goods_type": "Gạo",
                "closing_date": "28/07/2026",
                "vessel": "Tàu A",
                "recipient": "Công ty B",
                "carrier": "Vận tải ABC",
            }
        ],
        sheet_name="T07 26",
    )
    qtbot.addWidget(dialog)
    dialog.table.selectRow(0)

    assert dialog.selected_row == 12
    assert dialog.selected_source_sheet == "T06 26"
    assert dialog.table.columnCount() == 8
    assert dialog.table.item(0, 7).text() == "Vận tải ABC"
    assert "Cột" not in [
        dialog.table.horizontalHeaderItem(column).text()
        for column in range(dialog.table.columnCount())
    ]


def test_decision_tables_sort_by_container_and_preserve_initial_order(qtbot) -> None:
    conflicts = ConflictResolutionDialog(
        [
            {
                "conflict_id": "b",
                "type": "TARGET_CELL_OCCUPIED",
                "container": "VSCU0000002",
                "carrier": "Vận tải B",
            },
            {
                "conflict_id": "a",
                "type": "TARGET_CELL_OCCUPIED",
                "container": "DRYU0000001",
                "carrier": "Vận tải A",
            },
        ]
    )
    picker = ManualRowPickerDialog(
        [
            {"row_number": 2, "container": "VSCU0000002"},
            {"row_number": 3, "container": "DRYU0000001"},
        ]
    )
    reposts = RepostSelectionDialog(
        [
            {"source_item_index": 0, "container": "VSCU0000002"},
            {"source_item_index": 1, "container": "DRYU0000001"},
        ]
    )
    new_rows = PaymentNewRowsDialog(
        [
            {"item_id": "b", "container": "VSCU0000002", "values": {}},
            {"item_id": "a", "container": "DRYU0000001", "values": {}},
        ]
    )
    dialogs = (conflicts, picker, reposts, new_rows)
    for dialog in dialogs:
        qtbot.addWidget(dialog)

    tables_and_columns = (
        (conflicts.table, 0),
        (picker.table, 2),
        (reposts.table, 2),
        (new_rows.table, 4),
    )
    for table, container_column in tables_and_columns:
        assert table.isSortingEnabled()
        assert table.horizontalHeader().sortIndicatorSection() == -1
        assert table.item(0, container_column).text() == "VSCU0000002"
        table.sortItems(container_column, Qt.SortOrder.AscendingOrder)
        assert table.item(0, container_column).text() == "DRYU0000001"

    carrier_column = conflicts.COLUMNS.index("Bên vận tải")
    assert conflicts.table.item(0, carrier_column).text() == "Vận tải A"
    assert conflicts.table.cellWidget(
        0, conflicts.COLUMNS.index("Cách xử lý")
    ) is conflicts._action_combos["a"]
