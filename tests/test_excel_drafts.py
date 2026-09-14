from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from app.database import Database
from app.repositories.excel_draft_repository import ExcelDraftRepository
from app.services.excel.drafts import ExcelDraftService
from app.services.excel.models import WorkbookFingerprint


def _plan(
    tmp_path: Path,
    conflict: object,
    *,
    source_hash: str = "a" * 64,
    batch_id: int | None = 42,
    target_name: str = "BK.xlsx",
    selected_sheet: str | None = "T01 26",
    run_id: int | None = 100,
) -> object:
    return SimpleNamespace(
        operation="posting",
        batch_id=batch_id,
        batch_hash=source_hash,
        batch_path=tmp_path / "ready.json",
        target_path=tmp_path / target_name,
        target_fingerprint=WorkbookFingerprint(10, 20, "b" * 64),
        selected_sheet=selected_sheet,
        source_kind="BANG_KE",
        conflicts=[conflict],
        run_id=run_id,
    )


def _occupied_conflict(*, current: int = 100) -> object:
    return SimpleNamespace(
        conflict_id="occupied-1",
        conflict_type="TARGET_CELL_OCCUPIED",
        container="ABCD1234567",
        bl="BL-01",
        sqt=1,
        fee="NH",
        amount=200,
        sheet_name="T01 26",
        target_row=2,
        target_column=26,
        target_cell="Z2",
        current_value=current,
        allowed_actions=("KEEP_EXISTING", "OVERWRITE", "SKIP"),
        row_candidates=[],
        details={},
    )


def test_draft_restores_only_when_conflict_snapshot_is_still_valid(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    plan = _plan(tmp_path, _occupied_conflict())
    resolution = {
        "occupied-1": {
            "conflict_id": "occupied-1",
            "action": "KEEP_EXISTING",
        }
    }

    context_key = drafts.save(plan, "posting", resolution)
    restored_key, restored = drafts.restore(plan, "posting")

    assert restored_key == context_key
    assert restored == resolution

    changed_plan = _plan(tmp_path, _occupied_conflict(current=999))
    assert drafts.restore(changed_plan, "posting")[1] == {}
    database.close()


def test_failed_and_completed_choices_both_remain_available(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    plan = _plan(tmp_path, _occupied_conflict())
    resolution = {
        "occupied-1": {
            "conflict_id": "occupied-1",
            "action": "OVERWRITE",
        }
    }
    context_key = drafts.save(plan, "posting", resolution)

    drafts.mark_failed(context_key, "temporary error")
    assert drafts.restore(plan, "posting")[1] == resolution

    drafts.mark_completed(context_key)
    restored = drafts.restore_with_info(plan, "posting")
    assert restored.resolutions == resolution
    assert restored.status == "SUCCEEDED"
    database.close()


def test_same_batch_keeps_choices_when_source_version_changes(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    plan = _plan(tmp_path, _occupied_conflict())
    drafts.save(
        plan,
        "posting",
        {"occupied-1": {"conflict_id": "occupied-1", "action": "SKIP"}},
    )

    replacement = _plan(
        tmp_path, _occupied_conflict(), source_hash="c" * 64
    )
    assert drafts.restore(replacement, "posting")[1] == {
        "occupied-1": {"conflict_id": "occupied-1", "action": "SKIP"}
    }
    database.close()


def test_different_source_file_does_not_reuse_old_choices(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    plan = _plan(tmp_path, _occupied_conflict(), batch_id=42)
    drafts.save(
        plan,
        "posting",
        {"occupied-1": {"conflict_id": "occupied-1", "action": "SKIP"}},
    )

    replacement = _plan(tmp_path, _occupied_conflict(), batch_id=43)
    restored = drafts.restore_with_info(replacement, "posting")

    assert not restored.found
    assert restored.resolutions == {}
    database.close()


def test_sheet_selection_does_not_change_source_file_history(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    initial = _plan(tmp_path, _occupied_conflict(), selected_sheet=None)
    resolution = {
        "occupied-1": {"conflict_id": "occupied-1", "action": "OVERWRITE"}
    }
    drafts.save(initial, "posting", resolution)

    refined = _plan(tmp_path, _occupied_conflict(), selected_sheet="T02 26")

    assert drafts.restore(refined, "posting")[1] == resolution
    database.close()


def test_posting_group_assignments_are_restored_only_for_current_groups_and_sheets(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    plan = _plan(tmp_path, _occupied_conflict(), selected_sheet=None)
    plan.source_kind = "ASSISTANT"
    plan.source_groups = [
        SimpleNamespace(
            group_id="DOC_001", source_document_id="SOURCE_A",
            target_sheet="T07 26"
        ),
        SimpleNamespace(
            group_id="DOC_002", source_document_id="SOURCE_B",
            target_sheet="T06 26"
        ),
    ]
    plan.split_document_ids = {"SOURCE_A"}
    plan.sheet_candidates = [
        SimpleNamespace(sheet_name="T07 26"),
        SimpleNamespace(sheet_name="T06 26"),
    ]
    drafts.save(plan, "posting", {})

    restored_plan = _plan(tmp_path, _occupied_conflict(), selected_sheet=None)
    restored_plan.source_kind = "ASSISTANT"
    restored_plan.source_groups = [
        SimpleNamespace(
            group_id="DOC_001", source_document_id="SOURCE_A", target_sheet=None
        ),
        SimpleNamespace(
            group_id="DOC_002", source_document_id="SOURCE_B", target_sheet=None
        ),
    ]
    restored_plan.split_document_ids = set()
    restored_plan.sheet_candidates = plan.sheet_candidates

    assert drafts.restore(restored_plan, "posting")[1]["group_target_sheets"] == {
        "DOC_001": "T07 26",
        "DOC_002": "T06 26",
    }
    assert drafts.restore(restored_plan, "posting")[1]["split_document_ids"] == [
        "SOURCE_A"
    ]

    restored_plan.sheet_candidates = [SimpleNamespace(sheet_name="T07 26")]
    assert drafts.restore(restored_plan, "posting")[1]["group_target_sheets"] == {
        "DOC_001": "T07 26"
    }
    database.close()


def test_different_target_keeps_history_but_does_not_replay_choices(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    initial = _plan(tmp_path, _occupied_conflict(), target_name="BK-A.xlsx")
    drafts.save(
        initial,
        "posting",
        {"occupied-1": {"conflict_id": "occupied-1", "action": "SKIP"}},
    )

    changed_target = _plan(
        tmp_path, _occupied_conflict(), target_name="BK-B.xlsx"
    )
    restored = drafts.restore_with_info(changed_target, "posting")

    assert restored.found
    assert not restored.target_compatible
    assert restored.resolutions == {}
    database.close()


def test_cancelled_choices_remain_available(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    plan = _plan(tmp_path, _occupied_conflict())
    resolution = {
        "occupied-1": {"conflict_id": "occupied-1", "action": "KEEP_EXISTING"}
    }
    context_key = drafts.save(plan, "posting", resolution)

    drafts.mark_cancelled(context_key)
    restored = drafts.restore_with_info(plan, "posting")

    assert restored.status == "CANCELLED"
    assert restored.resolutions == resolution
    database.close()


def test_same_source_file_keeps_separate_history_for_each_operation(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    plan = _plan(tmp_path, _occupied_conflict(), batch_id=None)
    drafts.save(
        plan,
        "sync",
        {"occupied-1": {"conflict_id": "occupied-1", "action": "SKIP"}},
    )

    payment = drafts.restore_with_info(plan, "payment_sync")

    assert not payment.found
    assert drafts.restore(plan, "sync")[1]["occupied-1"]["action"] == "SKIP"
    database.close()


def test_new_run_keeps_choices_from_later_refine_stages(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    first_conflict = _occupied_conflict(current=100)
    first_plan = _plan(tmp_path, first_conflict, run_id=100)
    first_resolution = {
        "occupied-1": {"conflict_id": "occupied-1", "action": "SKIP"}
    }
    drafts.save(first_plan, "posting", first_resolution)

    later_conflict = _occupied_conflict(current=200)
    later_conflict.conflict_id = "later-conflict"
    later_plan = _plan(tmp_path, later_conflict, run_id=100)
    later_resolution = {
        "later-conflict": {
            "conflict_id": "later-conflict",
            "action": "KEEP_EXISTING",
        }
    }
    drafts.save(later_plan, "posting", later_resolution)

    next_run_first_plan = _plan(tmp_path, first_conflict, run_id=101)
    drafts.save(next_run_first_plan, "posting", first_resolution)
    next_run_later_plan = _plan(tmp_path, later_conflict, run_id=101)

    assert drafts.restore(next_run_later_plan, "posting")[1] == later_resolution
    database.close()


def test_cancel_action_is_not_replayed(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    plan = _plan(tmp_path, _occupied_conflict())
    plan.conflicts[0].allowed_actions = ("CANCEL_ALL", "SKIP")
    drafts.save(
        plan,
        "posting",
        {
            "occupied-1": {
                "conflict_id": "occupied-1",
                "action": "CANCEL_ALL",
            }
        },
    )

    assert drafts.restore(plan, "posting")[1] == {}
    database.close()


def test_dynamic_conflict_snapshot_restores_per_row_after_fingerprint_changes(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    base = _plan(tmp_path, _occupied_conflict(current=100))
    dynamic = _occupied_conflict(current=200)
    dynamic.conflict_id = "dynamic-invoice"
    resolution = {
        dynamic.conflict_id: {
            "conflict_id": dynamic.conflict_id,
            "action": "OVERWRITE",
        }
    }

    drafts.save(
        base,
        "posting",
        resolution,
        conflicts=[dynamic],
    )
    assert drafts.restore_for_conflicts(base, "posting", [dynamic]) == resolution

    changed = _plan(tmp_path, _occupied_conflict(current=100))
    changed.target_fingerprint = WorkbookFingerprint(10, 20, "c" * 64)
    restored = drafts.restore_with_info(changed, "posting")
    assert restored.target_changed
    assert drafts.restore_for_conflicts(changed, "posting", [dynamic]) == resolution

    changed_dynamic = _occupied_conflict(current=999)
    changed_dynamic.conflict_id = "dynamic-invoice"
    assert drafts.restore_for_conflicts(
        changed, "posting", [changed_dynamic]
    ) == {}
    database.close()


def test_clear_removes_only_the_current_file_choices(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    first = _plan(tmp_path, _occupied_conflict(), batch_id=42)
    second = _plan(tmp_path, _occupied_conflict(), batch_id=43)
    choice = {"occupied-1": {"conflict_id": "occupied-1", "action": "SKIP"}}
    drafts.save(first, "posting", choice)
    drafts.save(second, "posting", choice)

    assert drafts.clear(first, "posting")
    assert drafts.restore(first, "posting")[1] == {}
    assert drafts.restore(second, "posting")[1] == choice
    database.close()


def test_intermediate_choices_round_trip_with_item_validation(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")
    drafts = ExcelDraftService(ExcelDraftRepository(database))
    plan = _plan(tmp_path, _occupied_conflict())
    plan.month_candidates = [SimpleNamespace(target_sheet="T01 26")]
    plan.source_target_sheets = {"Tháng 1": None}
    plan.previously_posted_items = [
        SimpleNamespace(source_item_index=7, container="CONT1", amount=100)
    ]
    plan.new_rows = [
        SimpleNamespace(
            item_id="new-1", sqt=1, container="CONT1", values={"fee": 100}
        )
    ]
    drafts.save(
        plan,
        "posting",
        {
            "selected_sheet": "T01 26",
            "selected_month": 1,
            "source_target_sheets": {"Tháng 1": "T01 26"},
            "repost_source_indices": [7],
            "selected_new_rows": ["new-1"],
        },
    )

    restored = drafts.restore(plan, "posting")[1]
    assert restored["selected_sheet"] == "T01 26"
    assert restored["selected_month"] == 1
    assert restored["source_target_sheets"] == {"Tháng 1": "T01 26"}
    assert restored["repost_source_indices"] == [7]
    assert restored["selected_new_rows"] == ["new-1"]

    plan.new_rows[0].values = {"fee": 999}
    assert drafts.restore(plan, "posting")[1]["selected_new_rows"] == []
    database.close()
