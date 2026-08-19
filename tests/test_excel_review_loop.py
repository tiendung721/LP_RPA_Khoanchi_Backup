from __future__ import annotations

from types import SimpleNamespace
from typing import Any

from PySide6.QtCore import Qt

from app.services.excel.review import (
    CorrectionIssue,
    CorrectionRequiredError,
    ReviewOutcome,
    SourceDataChangedError,
    validate_conflict_resolutions,
)
from app.ui.excel_dialogs import ConflictResolutionDialog
from app.ui.excel_task_controller import ExcelTaskController
from app.ui.main_window import MainWindow, _ExcelRetryContext


def _conflict(conflict_id: str, *, message: str | None = None) -> dict[str, Any]:
    return {
        "conflict_id": conflict_id,
        "type": "TARGET_CELL_OCCUPIED",
        "message": message or f"Vấn đề {conflict_id}",
        "allowed_actions": ["KEEP_EXISTING", "OVERWRITE"],
        "default_action": "KEEP_EXISTING",
    }


def test_resolution_validation_collects_all_invalid_rows() -> None:
    conflicts = [
        {
            "conflict_id": "row",
            "type": "CONTAINER_NOT_FOUND",
            "allowed_actions": ["SELECT_ROW", "SKIP"],
            "row_candidates": [{"row": 8}],
        },
        {
            "conflict_id": "invoice",
            "type": "MULTIPLE_SOURCE_INVOICES",
            "allowed_actions": ["SELECT_INVOICE"],
            "details": {"invoice_candidates": ["HD-01"]},
        },
    ]

    issues = validate_conflict_resolutions(
        conflicts,
        {
            "row": {"action": "SELECT_ROW", "selected_row": 99},
            "invoice": {
                "action": "SELECT_INVOICE",
                "selected_invoice": "HD-XX",
            },
        },
    )

    assert {issue.conflict_ids[0] for issue in issues} == {"row", "invoice"}
    assert all(issue.message for issue in issues)


def test_conflict_dialog_reuses_choices_and_marks_only_latest_issues(qtbot) -> None:
    dialog = ConflictResolutionDialog(
        [_conflict("a"), _conflict("b")],
        initial_resolutions={"a": {"action": "OVERWRITE"}},
    )
    qtbot.addWidget(dialog)

    dialog.set_review(
        [_conflict("a"), _conflict("b"), _conflict("c")],
        issues=[
            CorrectionIssue(
                issue_id="invalid:b",
                message="Dòng B chưa tương thích.",
                conflict_ids=("b",),
            )
        ],
    )

    assert dialog.resolution_map()["a"]["action"] == "OVERWRITE"
    assert dialog.table.rowCount() == 3
    marked_rows: list[str] = []
    for row in range(dialog.table.rowCount()):
        item = dialog.table.item(row, 0)
        conflict_id = str(item.data(Qt.ItemDataRole.UserRole + 1))
        widget = dialog.table.cellWidget(row, dialog.COLUMNS.index("Cách xử lý"))
        if widget is not None and "#fff7d6" in widget.styleSheet():
            marked_rows.append(conflict_id)
    assert marked_rows == ["b"]
    issue_column = dialog.COLUMNS.index("Vấn đề")
    issue_text = next(
        dialog.table.item(row, issue_column).text()
        for row in range(dialog.table.rowCount())
        if dialog.table.item(row, 0).data(Qt.ItemDataRole.UserRole + 1) == "b"
    )
    assert issue_text == "Dòng B chưa tương thích."

    # Editing does not clear the last validation result.
    combo = dialog._action_combos["b"]
    combo.setCurrentIndex(combo.findData("OVERWRITE"))
    marked = next(
        dialog.table.item(row, 0)
        for row in range(dialog.table.rowCount())
        if dialog.table.item(row, 0).data(Qt.ItemDataRole.UserRole + 1) == "b"
    )
    marked_widget = dialog.table.cellWidget(
        marked.row(), dialog.COLUMNS.index("Cách xử lý")
    )
    assert marked_widget is not None
    assert "#fff7d6" in marked_widget.styleSheet()


def test_controller_shows_only_new_conflicts_after_json_row_selection(qtbot) -> None:
    source_choice = {
        "conflict_id": "source-choice",
        "type": "MULTIPLE_EXPENSE_SAME_CELL",
        "message": "Chọn một dòng JSON để ghi.",
        "allowed_actions": ["SELECT_SOURCE_ITEM"],
        "details": {
            "source_item_options": [
                {"source_item_index": 3},
                {"source_item_index": 7},
            ]
        },
    }
    base = {
        "operation": "posting",
        "has_changes": True,
        "conflicts": [source_choice],
    }

    class Service:
        def __init__(self) -> None:
            self.apply_count = 0
            self.refine_resolution_ids: list[set[str]] = []

        @staticmethod
        def analyze(*, progress_callback) -> Any:
            return base

        def refine(self, plan: Any, resolutions: Any, *, progress_callback) -> Any:
            self.refine_resolution_ids.append(set(resolutions))
            ids = {item["conflict_id"] for item in plan.get("conflicts", [])}
            if ids == {"source-choice"}:
                return {
                    "operation": "posting",
                    "has_changes": True,
                    "conflicts": [_conflict("dynamic")],
                }
            return {
                "operation": "posting",
                "has_changes": True,
                "conflicts": [],
            }

        def apply(self, plan: Any, resolutions: Any, *, progress_callback) -> Any:
            self.apply_count += 1
            return {"operation": "posting", "message": "Hoàn tất"}

    service = Service()
    controller = ExcelTaskController(expense_posting_service=service)
    ready: list[Any] = []
    corrections: list[ReviewOutcome] = []
    completed: list[Any] = []
    failed: list[Any] = []
    finished: list[str] = []
    controller.analysis_ready.connect(ready.append)
    controller.correction_required.connect(corrections.append)
    controller.completed.connect(completed.append)
    controller.failed.connect(failed.append)
    controller.finished.connect(finished.append)

    try:
        controller.start_posting()
        qtbot.waitUntil(lambda: bool(ready), timeout=2000)
        controller.prepare_plan(
            base,
            {
                "source-choice": {
                    "action": "SELECT_SOURCE_ITEM",
                    "selected_source_item_index": 7,
                }
            },
            operation="posting",
        )
        qtbot.waitUntil(lambda: bool(corrections), timeout=2000)

        assert controller.phase == "waiting_user"
        assert not failed
        assert not finished
        assert [
            issue.conflict_ids[0] for issue in corrections[-1].issues
        ] == ["dynamic"]
        assert [
            conflict["conflict_id"] for conflict in corrections[-1].conflicts
        ] == ["dynamic"]

        controller.prepare_plan(
            base,
            {
                "source-choice": {
                    "action": "SELECT_SOURCE_ITEM",
                    "selected_source_item_index": 7,
                },
                "dynamic": {"action": "OVERWRITE"},
            },
            operation="posting",
        )
        qtbot.waitUntil(lambda: bool(completed), timeout=2000)

        assert service.apply_count == 1
        assert service.refine_resolution_ids[-2:] == [
            {"source-choice"},
            {"dynamic"},
        ]
        assert finished == ["posting"]
        assert not controller.is_busy
    finally:
        controller.shutdown()


def test_controller_keeps_session_when_apply_preflight_is_correctable(qtbot) -> None:
    base = {
        "operation": "sync",
        "has_changes": True,
        "conflicts": [_conflict("row")],
    }

    class Service:
        should_fail = True

        @staticmethod
        def analyze(*, progress_callback) -> Any:
            return base

        def apply(self, plan: Any, resolutions: Any, *, progress_callback) -> Any:
            if self.should_fail:
                raise CorrectionRequiredError(
                    [
                        CorrectionIssue(
                            issue_id="preflight:row",
                            message="Dòng vẫn chưa tương thích.",
                            conflict_ids=("row",),
                        )
                    ]
                )
            return {"operation": "sync", "message": "Hoàn tất"}

    service = Service()
    controller = ExcelTaskController(daily_sync_service=service)
    ready: list[Any] = []
    corrections: list[ReviewOutcome] = []
    completed: list[Any] = []
    failed: list[Any] = []
    finished: list[str] = []
    controller.analysis_ready.connect(ready.append)
    controller.correction_required.connect(corrections.append)
    controller.completed.connect(completed.append)
    controller.failed.connect(failed.append)
    controller.finished.connect(finished.append)

    try:
        controller.start_sync()
        qtbot.waitUntil(lambda: bool(ready), timeout=2000)
        resolution = {"row": {"action": "KEEP_EXISTING"}}
        controller.apply_plan(base, resolution, operation="sync")
        qtbot.waitUntil(lambda: bool(corrections), timeout=2000)

        assert controller.phase == "waiting_user"
        assert not failed
        assert not finished

        service.should_fail = False
        controller.apply_plan(base, resolution, operation="sync")
        qtbot.waitUntil(lambda: bool(completed), timeout=2000)
        assert finished == ["sync"]
    finally:
        controller.shutdown()


def test_source_reload_context_keeps_posting_sheet_and_repost_selection() -> None:
    owner = SimpleNamespace(
        _excel_review_plan=SimpleNamespace(
            batch_id=45,
            selected_sheet="T07 26",
            repost_selection_done=True,
            repost_source_indices={4, 1},
        )
    )

    context = MainWindow._excel_retry_context_for(
        owner,
        SourceDataChangedError("JSON nguồn"),
    )

    assert context == _ExcelRetryContext(
        operation="posting",
        batch_id=45,
        sheet_name="T07 26",
        repost_source_indices=(1, 4),
        highlight_unresolved=True,
    )


def test_source_reload_restarts_posting_without_asking_for_month_again() -> None:
    calls: list[dict[str, Any]] = []

    class Tasks:
        @staticmethod
        def start_posting(**kwargs: Any) -> None:
            calls.append(kwargs)

    owner = SimpleNamespace(
        _excel_tasks=Tasks(),
        _excel_context=None,
        _excel_reanalysis_from_source_change=False,
    )
    context = _ExcelRetryContext(
        operation="posting",
        batch_id=45,
        sheet_name="T07 26",
        repost_source_indices=(1, 4),
        highlight_unresolved=True,
    )

    MainWindow._restart_expense_posting(owner, context)

    assert calls == [
        {
            "batch_id": 45,
            "sheet_name": "T07 26",
            "repost_source_indices": (1, 4),
        }
    ]
    assert owner._excel_context == "workflow"
    assert owner._excel_reanalysis_from_source_change is True
