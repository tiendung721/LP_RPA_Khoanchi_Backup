from __future__ import annotations

import sqlite3
from pathlib import Path

from app.constants import SQLITE_SCHEMA_VERSION
from app.database import Database
from app.repositories.batch_repository import BatchRepository
from app.repositories.excel_run_repository import ExcelRunRepository
from app.repositories.expense_posting_repository import ExpensePostingRepository


def _create_batch(database: Database, tmp_path: Path, *, sha: str) -> int:
    batch = BatchRepository(database).create_batch(
        source_filename="ket_qua_boc_tach.json",
        source_output_path=tmp_path / "Output" / "ket_qua_boc_tach.json",
        original_archive_path=tmp_path / "Archive" / "original.json",
        working_path=tmp_path / "Workspace" / "working.json",
        ready_path=tmp_path / "Ready" / "ready.json",
        sha256=sha,
    )
    return batch.id


def test_v2_database_is_migrated_transactionally_to_excel_schema(
    tmp_path: Path,
) -> None:
    path = tmp_path / "legacy-v2.db"
    connection = sqlite3.connect(path)
    connection.row_factory = sqlite3.Row
    connection.execute("BEGIN")
    Database._migration_1(connection)
    Database._migration_2(connection)
    connection.execute("PRAGMA user_version = 2")
    connection.commit()
    connection.close()

    database = Database(path)

    assert database.query_one("PRAGMA user_version")[0] == SQLITE_SCHEMA_VERSION
    assert database.query_one(
        "SELECT name FROM sqlite_master WHERE name = 'excel_runs'"
    )
    excel_run_columns = {
        row["name"] for row in database.query_all("PRAGMA table_info(excel_runs)")
    }
    assert "item_outcomes" in excel_run_columns
    assert database.query_one(
        "SELECT name FROM sqlite_master WHERE name = 'expense_posting_items'"
    )
    assert database.query_one(
        "SELECT name FROM sqlite_master WHERE name = 'expense_carry_forwards'"
    )
    history_index = database.query_one(
        """
        SELECT sql FROM sqlite_master
        WHERE type = 'index' AND name = 'idx_expense_posting_success_source'
        """
    )
    assert history_index is not None
    assert "WHERE status IN ('POSTED', 'ALREADY_EXISTS')" in history_index["sql"]
    assert database.query_one(
        """
        SELECT name FROM sqlite_master
        WHERE type = 'index' AND name = 'ux_expense_posting_success_source'
        """
    ) is None
    columns = {
        row["name"]
        for row in database.query_all("PRAGMA table_info(expense_posting_items)")
    }
    assert {
        "invoice_no",
        "invoice_selected",
        "invoice_target_column",
        "invoice_target_cell",
        "invoice_value_before",
        "invoice_value_after",
        "invoice_action",
    }.issubset(columns)
    database.close()


def test_excel_run_repository_tracks_fingerprints_counts_and_latest_sheet(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "app_state.db")
    repository = ExcelRunRepository(database)
    before = {"size": 120, "mtime_ns": 456, "sha256": "a" * 64}

    run = repository.create_run(
        operation="DAILY_SYNC",
        source_path=tmp_path / "daily.xlsx",
        target_path=tmp_path / "bk.xlsx",
        source_fingerprint={"sha256": "b" * 64},
        target_fingerprint_before=before,
    )
    completed = repository.finish_run(
        run.id,
        status="SUCCEEDED",
        sheet_name="T07 26",
        backup_path=tmp_path / "Backup" / "bk.xlsx",
        target_fingerprint_after={"sha256": "c" * 64},
        total_items=3,
        changed_items=2,
        skipped_items=1,
        conflict_count=1,
        item_outcomes=[
            {
                "item_id": "daily-2",
                "status": "WRITTEN",
                "fields": [{"field_name": "SQT", "target_value_after": 101}],
            }
        ],
    )

    assert completed.target_fingerprint_before == before
    assert completed.target_fingerprint_after == {"sha256": "c" * 64}
    assert completed.backup_path == tmp_path / "Backup" / "bk.xlsx"
    assert completed.completed_at is not None
    assert completed.item_outcomes[0]["item_id"] == "daily-2"
    assert repository.get_latest_sync_sheet() == "T07 26"
    assert repository.get_latest(
        operation="DAILY_SYNC",
        statuses={"SUCCEEDED"},
    ) == completed
    database.close()


def test_posting_repository_round_trips_items_and_filters_successful_sources(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "app_state.db")
    batch_id = _create_batch(database, tmp_path, sha="d" * 64)
    run = ExcelRunRepository(database).create_run(operation="EXPENSE_POSTING")
    repository = ExpensePostingRepository(database)

    records = repository.create_items(
        [
            {
                "source_item_index": 0,
                "container": "ABCD1234567",
                "bl": None,
                "match_reason": "CONTAINER_ONLY",
                "fee_original": "CB",
                "fee_selected": "CB",
                "rule": "HD",
                "amount": 1_250_000,
                "sheet_name": "T07 26",
                "target_row": 42,
                "target_column": 17,
                "target_cell": "Q42",
                "value_before": {"kind": "empty", "value": None},
                "value_after": 1_250_000,
                "action": "WRITE",
                "invoice_no": "INV-001",
                "invoice_selected": "INV-001",
                "invoice_target_column": 18,
                "invoice_target_cell": "R42",
                "invoice_value_before": None,
                "invoice_value_after": "INV-001",
                "invoice_action": "OVERWRITE",
                "carrier_source": "NHS",
                "carrier_effective": "NHS / TP",
                "carrier_group": "NAM",
                "carrier_target_column": 23,
                "carrier_target_cell": "W42",
                "carrier_value_before": "TP",
                "carrier_value_after": "NHS / TP",
                "carrier_action": "APPEND_CARRIER",
                "status": "POSTED",
            },
            {
                "source_item_index": 1,
                "container": "EFGH1234567",
                "fee_original": "NV",
                "amount": 500_000,
                "status": "ALREADY_EXISTS",
                "value_before": 500_000,
                "value_after": 500_000,
            },
        ],
        run_id=run.id,
        batch_id=batch_id,
        batch_hash="e" * 64,
    )

    assert len(records) == 2
    assert records[0].target_cell == "Q42"
    assert records[0].match_reason == "CONTAINER_ONLY"
    assert records[0].value_before == {"kind": "empty", "value": None}
    assert records[0].invoice_no == "INV-001"
    assert records[0].invoice_selected == "INV-001"
    assert records[0].invoice_target_cell == "R42"
    assert records[0].invoice_value_after == "INV-001"
    assert records[0].invoice_action == "OVERWRITE"
    assert records[0].carrier_source == "NHS"
    assert records[0].carrier_effective == "NHS / TP"
    assert records[0].carrier_group == "NAM"
    assert records[0].carrier_target_column == 23
    assert records[0].carrier_target_cell == "W42"
    assert records[0].carrier_value_before == "TP"
    assert records[0].carrier_value_after == "NHS / TP"
    assert records[0].carrier_action == "APPEND_CARRIER"
    assert repository.successful_source_indices("e" * 64) == {0, 1}
    assert repository.batch_has_successful_items("e" * 64)
    assert repository.is_source_item_posted("e" * 64, 1)
    assert repository.get_latest_successful_for_batch("e" * 64) is not None
    database.close()


def test_posting_history_allows_multiple_successful_reposts(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "app_state.db")
    batch_id = _create_batch(database, tmp_path, sha="f" * 64)
    runs = ExcelRunRepository(database)
    postings = ExpensePostingRepository(database)
    first_run = runs.create_run(operation="EXPENSE_POSTING")
    second_run = runs.create_run(operation="EXPENSE_POSTING")
    common = {
        "batch_id": batch_id,
        "batch_hash": "1" * 64,
        "source_item_index": 4,
        "fee_original": "HH",
        "amount": 100_000,
    }

    postings.create_item(run_id=first_run.id, status="POSTED", **common)
    retry = postings.create_item(run_id=second_run.id, status="PLANNED", **common)

    postings.update_item(retry.id, status="ALREADY_EXISTS")

    assert postings.require_by_id(retry.id).status == "ALREADY_EXISTS"
    assert postings.successful_source_indices("1" * 64) == {4}
    latest = postings.latest_successful_items("1" * 64)
    assert [item.id for item in latest] == [retry.id]
    database.close()


def test_posting_repository_round_trips_carry_forward_mapping(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "app_state.db")
    repository = ExpensePostingRepository(database)
    workbook = tmp_path / "BK 2026.xlsx"

    created = repository.save_carry_forward(
        workbook_path=workbook,
        source_sheet="T06 26",
        source_row=28,
        source_sqt=619,
        container="DRYU3026167",
        source_signature="a" * 64,
        target_sheet="T07 26",
        target_row=40,
    )
    loaded = repository.get_carry_forward(
        workbook_path=workbook,
        source_sheet="T06 26",
        source_row=28,
        source_sqt=619,
        container="DRYU3026167",
        target_sheet="T07 26",
    )

    assert loaded == created
    assert loaded is not None
    assert loaded.target_row == 40
    updated = repository.save_carry_forward(
        workbook_path=workbook,
        source_sheet="T06 26",
        source_row=28,
        source_sqt=619,
        container="DRYU3026167",
        source_signature="b" * 64,
        target_sheet="T07 26",
        target_row=41,
    )
    assert updated.id == created.id
    assert updated.target_row == 41
    assert updated.source_signature == "b" * 64
    database.close()
