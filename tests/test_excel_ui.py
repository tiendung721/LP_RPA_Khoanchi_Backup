from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QMessageBox,
    QPushButton,
)

import app.ui.main_window as main_window_module
from app.ui.excel_dialogs import (
    ConflictResolutionDialog,
    ManualRowPickerDialog,
    MonthSelectionDialog,
    RepostSelectionDialog,
)
from app.ui.excel_task_controller import ExcelTaskController
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
    assert page.sync_daily_button.text() == "Đồng bộ dữ liệu Hàng ngày"
    assert page.post_expenses_button.text() == "Nhập khoản chi vào BK"
    assert page.sync_payment_button.text() == "Đồng bộ BK → Thanh toán"
    assert page.sync_status_label.text() == "Đồng bộ gần nhất: —"
    assert page.posting_status_label.text() == "Nhập khoản chi gần nhất: —"
    assert (
        page.payment_sync_status_label.text()
        == "Đồng bộ BK → Thanh toán gần nhất: —"
    )


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
    assert page.sync_daily_button.text() == "Đồng bộ dữ liệu Hàng ngày"


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
            self.selected_sheet_name = "Tháng 8"

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
    assert tasks.sync_calls == [{"source_sheet_name": "Tháng 8"}]


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

    captured: dict[str, Any] = {}

    def question(
        _parent: Any,
        title: str,
        text: str,
        *_args: Any,
    ) -> QMessageBox.StandardButton:
        captured["title"] = title
        captured["text"] = text
        return QMessageBox.StandardButton.Yes

    monkeypatch.setattr(QMessageBox, "question", question)
    tasks = Tasks()
    owner = SimpleNamespace(
        _excel_tasks=tasks,
        _excel_operation="sync",
        _show_excel_error=lambda *_args, **_kwargs: None,
    )

    MainWindow._excel_analysis_ready(owner, plan)

    assert len(tasks.apply_calls) == 1
    assert tasks.cancel_calls == 0
    assert captured["title"] == "Xác nhận đồng bộ toàn sheet"
    assert "Cập nhật: 5 dòng" in captured["text"]
    assert "Thêm mới: 2 dòng" in captured["text"]
    assert "Chỉ có ở BK, được giữ lại: 1 dòng" in captured["text"]
    assert "Thiếu SQT, được bỏ qua: 3 dòng" in captured["text"]


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
        def __init__(self, _conflicts: Any, _parent: Any) -> None:
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

    captured: dict[str, str] = {}

    def question(
        _parent: Any,
        title: str,
        text: str,
        *_args: Any,
    ) -> QMessageBox.StandardButton:
        captured["title"] = title
        captured["text"] = text
        return QMessageBox.StandardButton.Yes

    monkeypatch.setattr(QMessageBox, "question", question)
    tasks = Tasks()
    owner = SimpleNamespace(_excel_tasks=tasks)

    MainWindow._handle_payment_sync_plan(owner, plan)

    assert len(tasks.apply_calls) == 1
    assert "T06 26 HP" in captured["text"]
    assert "T05 26 HP" in captured["text"]
    assert "T06 26 NAM" in captured["text"]
    assert "T05 26 NAM" in captured["text"]
    assert captured["title"] == "Xác nhận đồng bộ BK → Thanh toán"
    assert "Sheet Thanh toán mới: Có, tạo từ T05 26" in captured["text"]


@pytest.mark.parametrize(
    ("operation", "result_fields", "expected_title", "expected_detail"),
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
            "Đồng bộ thành công",
            "Sheet: T07 26",
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
            "Nhập khoản chi hoàn tất",
            "Đã nhập: 10 khoản",
        ),
    ),
)
def test_bk_completion_can_open_the_written_bk_file(
    monkeypatch,
    tmp_path: Path,
    operation: str,
    result_fields: dict[str, Any],
    expected_title: str,
    expected_detail: str,
) -> None:
    bk_path = tmp_path / "BK 2026.xlsx"
    bk_path.touch()
    opened: list[Path] = []

    class CompletionMessage:
        class Icon:
            Information = object()

        class ButtonRole:
            ActionRole = object()

        class StandardButton:
            Ok = object()

        instance: Any = None

        def __init__(self, _parent: Any) -> None:
            CompletionMessage.instance = self
            self.title = ""
            self.text = ""
            self.open_button = None
            self.clicked = None

        def setIcon(self, _icon: Any) -> None:
            pass

        def setWindowTitle(self, title: str) -> None:
            self.title = title

        def setText(self, text: str) -> None:
            self.text = text

        def addButton(self, button: Any, _role: Any = None) -> Any:
            if button == "Mở file BK":
                self.open_button = object()
                return self.open_button
            return object()

        def exec(self) -> None:
            self.clicked = self.open_button

        def clickedButton(self) -> Any:
            return self.clicked

    class Tasks:
        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return operation

    monkeypatch.setattr(main_window_module, "QMessageBox", CompletionMessage)
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

    message = CompletionMessage.instance
    assert message.title == expected_title
    assert message.open_button is not None
    assert expected_detail in message.text
    assert opened == [bk_path]


def test_payment_completion_displays_carrier_invoice_summary(
    monkeypatch,
    tmp_path: Path,
) -> None:
    payment_path = tmp_path / "Thanh toan.xlsm"
    payment_path.touch()
    opened: list[Path] = []

    class CompletionMessage:
        class Icon:
            Information = object()

        class ButtonRole:
            ActionRole = object()

        class StandardButton:
            Ok = object()

        instance: Any = None

        def __init__(self, _parent: Any) -> None:
            CompletionMessage.instance = self
            self.text = ""
            self.open_button = None

        def setIcon(self, _icon: Any) -> None:
            pass

        def setWindowTitle(self, _title: str) -> None:
            pass

        def setText(self, text: str) -> None:
            self.text = text

        def addButton(self, button: Any, _role: Any = None) -> Any:
            if button == "Mở file Thanh toán":
                self.open_button = object()
                return self.open_button
            return object()

        def exec(self) -> None:
            pass

        def clickedButton(self) -> Any:
            return self.open_button

    class Tasks:
        @staticmethod
        def normalize_operation(_operation: Any) -> str:
            return "payment_sync"

    monkeypatch.setattr(main_window_module, "QMessageBox", CompletionMessage)
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

    assert "Tổng hợp theo hóa đơn và bên vận tải" in CompletionMessage.instance.text
    assert "Số hóa đơn: 3" in CompletionMessage.instance.text
    assert "Tổng tiền kỳ: 1.700.000" in CompletionMessage.instance.text
    assert "Hóa đơn thuộc nhiều bên: 1" in CompletionMessage.instance.text
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

    assert duplicate_labels == ["Bỏ qua", "Chọn dòng"]
    assert occupied_labels == ["Giữ nguyên", "Ghi đè", "Bỏ qua"]
    assert "repeated" in dialog._selector_buttons
    assert not any(
        "_" in label
        for label in duplicate_labels + occupied_labels
    )


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
            }
        ],
        sheet_name="T07 26",
    )
    qtbot.addWidget(dialog)
    dialog.table.selectRow(0)

    assert dialog.selected_row == 12
    assert dialog.selected_source_sheet == "T06 26"
    assert dialog.table.columnCount() == 7
    assert "Cột" not in [
        dialog.table.horizontalHeaderItem(column).text()
        for column in range(dialog.table.columnCount())
    ]
