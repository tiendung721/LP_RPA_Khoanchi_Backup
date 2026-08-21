from __future__ import annotations

from pathlib import Path

import json
import sqlite3

from app.config import (
    AppPaths,
    AppSettings,
    ConfigManager,
    default_data_root,
    migrate_legacy_runtime_layout,
    software_root,
)
from app.constants import SQLITE_SCHEMA_VERSION
from app.database import Database
from app.models import BatchStatus
from app.repositories.batch_repository import BatchRepository


def test_first_load_creates_settings_and_runtime_directories(tmp_path: Path) -> None:
    manager = ConfigManager(tmp_path / "runtime")

    settings = manager.load()

    assert manager.settings_path.is_file()
    assert settings.data_root == tmp_path / "runtime"
    assert all(path.is_dir() for path in settings.paths.directories)


def test_default_data_root_is_the_portable_application_directory() -> None:
    assert default_data_root() == software_root()


def test_runtime_files_are_grouped_below_output_system(tmp_path: Path) -> None:
    paths = AppPaths.from_data_root(tmp_path / "runtime", tmp_path / "Output")

    assert paths.system_dir == tmp_path / "Output" / "_system"
    assert paths.archive_original_dir == (
        tmp_path / "Output" / "_system" / "Archive" / "Original"
    )
    assert paths.workspace_dir == tmp_path / "Output" / "_system" / "Workspace"
    assert paths.ready_dir == tmp_path / "Output" / "_system" / "Ready"
    assert paths.rejected_dir == tmp_path / "Output" / "_system" / "Rejected"
    assert paths.excel_temp_dir == tmp_path / "Output" / "_system" / "Excel" / "Temp"
    assert paths.excel_backup_dir == (
        tmp_path / "Output" / "_system" / "Excel" / "Backup"
    )
    assert paths.excel_reports_dir == (
        tmp_path / "Output" / "_system" / "Excel" / "Reports"
    )
    assert paths.rpa_dir == tmp_path / "Output" / "_system" / "RPA"


def test_legacy_runtime_layout_is_moved_without_changing_files(
    tmp_path: Path,
) -> None:
    root = tmp_path / "runtime"
    paths = AppPaths.from_data_root(root, tmp_path / "Output")
    legacy_files = {
        root / "Archive" / "Original" / "original.json": b"original",
        root / "Workspace" / "1" / "working.json": b"working",
        root / "Ready" / "ready.json": b"ready",
        root / "Rejected" / "bad.json": b"bad",
    }
    for path, content in legacy_files.items():
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(content)
    paths.ensure_directories()

    assert migrate_legacy_runtime_layout(paths) == 4

    assert not (root / "Archive").exists()
    assert not (root / "Workspace").exists()
    assert not (root / "Ready").exists()
    assert not (root / "Rejected").exists()
    assert (paths.archive_original_dir / "original.json").read_bytes() == b"original"
    assert (paths.workspace_dir / "1" / "working.json").read_bytes() == b"working"
    assert (paths.ready_dir / "ready.json").read_bytes() == b"ready"
    assert (paths.rejected_dir / "bad.json").read_bytes() == b"bad"


def test_settings_round_trip_utf8(tmp_path: Path) -> None:
    root = tmp_path / "Dữ liệu quyết toán"
    paths = AppPaths.from_data_root(root, tmp_path / "Kết quả")
    manager = ConfigManager(paths=paths)
    settings = AppSettings(
        data_root=root,
        assistant_bat_path=str(root / "Mở trợ lý.bat"),
        bang_ke_assistant_bat_path=str(root / "Mở tool bảng kê.bat"),
        output_dir=tmp_path / "Kết quả",
        daily_workbook_path=root / "Hàng ngày 2026.xlsx",
        bk_workbook_path=root / "BK Tổng hợp 2026.xlsm",
        rpa_expense_bat_path=str(root / "Chạy PAD quyết toán.bat"),
    )

    manager.save(settings)
    loaded = manager.load()

    assert loaded.assistant_bat_path == settings.assistant_bat_path
    assert loaded.bang_ke_assistant_bat_path == settings.bang_ke_assistant_bat_path
    assert loaded.output_dir == settings.output_dir
    assert loaded.daily_workbook_path == settings.daily_workbook_path
    assert loaded.bk_workbook_path == settings.bk_workbook_path
    assert loaded.rpa_expense_bat_path == settings.rpa_expense_bat_path
    assert manager.settings_path.read_bytes().startswith(b"{")


def test_database_migration_and_active_batch_restore(tmp_path: Path) -> None:
    database = Database(tmp_path / "Database" / "app_state.db")
    repository = BatchRepository(database)
    working = tmp_path / "Workspace" / "1" / "ket_qua_boc_tach.json"
    working.parent.mkdir(parents=True)
    working.write_text('{"v":1,"d":[]}', encoding="utf-8")
    original = tmp_path / "Archive" / "original.json"
    original.parent.mkdir(parents=True)
    original.write_text('{"v":1,"d":[]}', encoding="utf-8")

    batch = repository.create_batch(
        source_filename="ket_qua_boc_tach.json",
        source_output_path=tmp_path / "Output" / "ket_qua_boc_tach.json",
        original_archive_path=original,
        working_path=working,
        sha256="a" * 64,
        status=BatchStatus.REVIEWING,
    )
    repository.set_active_batch_id(batch.id)

    restored = repository.restore_active_batch()

    assert restored is not None
    assert restored.id == batch.id
    assert repository.get_app_state("active_batch_id") == str(batch.id)
    group_columns = {
        str(row[1])
        for row in database.query_all(
            "PRAGMA table_info(sea_freight_reconciliation_groups)"
        )
    }
    contribution_columns = {
        str(row[1])
        for row in database.query_all(
            "PRAGMA table_info(sea_freight_invoice_contributions)"
        )
    }
    assert {
        "reconciliation_month", "reconciliation_year", "bk_issue",
        "bk_invalid_container_count", "bk_duplicate_container_count",
        "revision_no", "supersedes_group_id", "is_current",
        "primary_source_batch_id", "primary_source_item_index",
    }.issubset(group_columns)
    assert {
        "source_kind", "source_document_id", "source_document_name"
    }.issubset(contribution_columns)
    assert database.query_one(
        "SELECT 1 FROM sqlite_master WHERE type = 'index' "
        "AND name = 'idx_sea_freight_contribution_invoice'"
    ) is not None
    database.close()


def test_legacy_settings_are_rewritten_without_browser_or_inbox_keys(
    tmp_path: Path,
) -> None:
    paths = AppPaths.from_data_root(tmp_path / "runtime", tmp_path / "Output")
    paths.ensure_directories()
    paths.settings_path.write_text(
        json.dumps(
            {
                "data_root": str(paths.data_root),
                "gpt_url": "https://example.invalid",
                "inbox_dir": str(tmp_path / "Inbox"),
                "browser_executable": "chrome.exe",
            }
        ),
        encoding="utf-8",
    )

    loaded = ConfigManager(paths=paths).load()
    rewritten = json.loads(paths.settings_path.read_text(encoding="utf-8"))

    assert loaded.assistant_bat_path == ""
    assert loaded.output_dir == paths.output_dir
    assert loaded.daily_workbook_path == ""
    assert loaded.bk_workbook_path == ""
    assert loaded.payment_workbook_path == ""
    assert set(rewritten) == {
        "data_root",
            "assistant_bat_path",
            "bang_ke_assistant_bat_path",
        "output_dir",
        "daily_workbook_path",
        "bk_workbook_path",
        "payment_workbook_path",
        "rpa_expense_bat_path",
    }


def test_database_v6_removes_legacy_pad_lookup_jobs(tmp_path: Path) -> None:
    path = tmp_path / "legacy-pad.db"
    connection = sqlite3.connect(path)
    connection.execute(
        "CREATE TABLE container_lookup_jobs "
        "(job_id TEXT PRIMARY KEY, requested_bl TEXT)"
    )
    connection.execute("PRAGMA user_version = 5")
    connection.commit()
    connection.close()

    database = Database(path)
    row = database.query_one(
        "SELECT name FROM sqlite_master "
        "WHERE type = 'table' AND name = 'container_lookup_jobs'"
    )

    assert row is None
    assert database.query_one("PRAGMA user_version")[0] == SQLITE_SCHEMA_VERSION
    database.close()


def test_database_creates_excel_resolution_draft_storage(tmp_path: Path) -> None:
    database = Database(tmp_path / "app.db")

    columns = {
        row["name"]
        for row in database.query_all(
            "PRAGMA table_info(excel_resolution_drafts)"
        )
    }

    assert {
        "context_key",
        "operation",
        "context_json",
        "payload_json",
        "status",
        "last_error",
    }.issubset(columns)
    assert database.query_one("PRAGMA user_version")[0] == SQLITE_SCHEMA_VERSION
    database.close()


def test_database_creates_latest_resolution_storage_by_source_file(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "app.db")

    columns = {
        row["name"]
        for row in database.query_all(
            "PRAGMA table_info(excel_resolution_latest)"
        )
    }

    assert {
        "source_file_key",
        "operation",
        "context_json",
        "payload_json",
        "status",
        "run_id",
        "last_error",
    }.issubset(columns)
    assert database.query_one("PRAGMA user_version")[0] == SQLITE_SCHEMA_VERSION
    database.close()


def test_database_migrates_legacy_resolution_without_reimporting_it(
    tmp_path: Path,
) -> None:
    path = tmp_path / "app.db"
    seed = Database(path)
    seed.execute("DROP TABLE excel_resolution_latest")
    seed.execute("DELETE FROM excel_resolution_drafts")
    context = {
        "operation": "posting",
        "batch_id": 43,
        "source_path": "c:/output/ready.json",
        "target_path": "c:/output/bk.xlsx",
    }
    seed.execute(
        """
        INSERT INTO excel_resolution_drafts (
            context_key, operation, context_json, payload_json, status,
            created_at, updated_at, completed_at, last_error
        ) VALUES (?, 'posting', ?, ?, 'FAILED', ?, ?, NULL, ?)
        """,
        (
            "legacy-key",
            json.dumps(context),
            json.dumps({"version": 1, "entries": {}, "globals": {}}),
            "2026-08-14T10:00:00+07:00",
            "2026-08-14T10:05:00+07:00",
            "temporary error",
        ),
    )
    seed.connection.execute("PRAGMA user_version = 14")
    seed.close()

    migrated = Database(path)
    row = migrated.query_one(
        "SELECT * FROM excel_resolution_latest WHERE source_file_key = ?",
        ("posting:batch:43",),
    )
    assert row is not None
    assert row["status"] == "FAILED"
    migrated.execute(
        "UPDATE excel_resolution_latest SET status = 'SUCCEEDED' "
        "WHERE source_file_key = ?",
        ("posting:batch:43",),
    )
    migrated.close()

    reopened = Database(path)
    assert reopened.query_one(
        "SELECT status FROM excel_resolution_latest WHERE source_file_key = ?",
        ("posting:batch:43",),
    )["status"] == "SUCCEEDED"
    reopened.close()


def test_copied_settings_rebase_paths_inside_the_old_bundle(tmp_path: Path) -> None:
    old_root = tmp_path / "old-bundle"
    new_root = tmp_path / "new-bundle"
    paths = AppPaths.from_data_root(new_root, new_root / "Output")
    paths.ensure_directories()
    paths.settings_path.write_text(
        json.dumps(
            {
                "data_root": str(old_root),
                "assistant_bat_path": str(tmp_path / "external" / "assistant.bat"),
                "output_dir": str(old_root / "Output"),
            }
        ),
        encoding="utf-8",
    )

    manager = ConfigManager(paths=paths)
    loaded = manager.load()

    assert manager.relocated_from == old_root
    assert loaded.data_root == new_root
    assert loaded.output_dir == new_root / "Output"
    assert loaded.assistant_bat_path == str(tmp_path / "external" / "assistant.bat")


def test_database_v1_renames_source_path_and_scan_state(tmp_path: Path) -> None:
    path = tmp_path / "legacy.db"
    connection = sqlite3.connect(path)
    connection.execute(
        "CREATE TABLE batches (id INTEGER PRIMARY KEY, source_inbox_path TEXT)"
    )
    connection.execute("CREATE TABLE app_state (key TEXT PRIMARY KEY, value TEXT)")
    connection.execute(
        "INSERT INTO app_state(key, value) VALUES ('last_inbox_scan_at', 'old-time')"
    )
    connection.execute("PRAGMA user_version = 1")
    connection.commit()
    connection.close()

    database = Database(path)
    columns = {
        row["name"] for row in database.query_all("PRAGMA table_info(batches)")
    }

    assert "source_output_path" in columns
    assert "source_inbox_path" not in columns
    assert database.query_one(
        "SELECT value FROM app_state WHERE key = 'last_output_scan_at'"
    )["value"] == "old-time"
    assert database.query_one(
        "SELECT value FROM app_state WHERE key = 'last_inbox_scan_at'"
    ) is None
    database.close()


def test_database_rebases_only_paths_inside_the_old_root(tmp_path: Path) -> None:
    old_root = tmp_path / "old-bundle"
    new_root = tmp_path / "new-bundle"
    external_output = tmp_path / "shared-output" / "ket_qua_boc_tach.json"
    database = Database(tmp_path / "app_state.db")
    repository = BatchRepository(database)
    batch = repository.create_batch(
        source_filename="ket_qua_boc_tach.json",
        source_output_path=external_output,
        original_archive_path=old_root / "Archive" / "Original" / "original.json",
        working_path=old_root / "Workspace" / "1" / "ket_qua_boc_tach.json",
        ready_path=old_root / "Ready" / "ready.json",
        sha256="b" * 64,
    )

    assert database.rebase_paths(old_root, new_root) == 1
    moved = repository.get_by_id(batch.id)

    assert moved is not None
    assert moved.source_output_path == external_output
    assert moved.original_archive_path == (
        new_root / "Archive" / "Original" / "original.json"
    )
    assert moved.working_path == (
        new_root / "Workspace" / "1" / "ket_qua_boc_tach.json"
    )
    assert moved.ready_path == new_root / "Ready" / "ready.json"
    database.close()


def test_database_v17_removes_vat_history_and_recalculates_aggregates(
    tmp_path: Path,
) -> None:
    path = tmp_path / "vat-history.db"
    database = Database(path)
    repository = BatchRepository(database)
    batch = repository.create_batch(
        source_filename="old-bang-ke.json",
        source_output_path=tmp_path / "old-bang-ke.json",
        original_archive_path=tmp_path / "old-bang-ke.json",
        working_path=tmp_path / "old-bang-ke.json",
        ready_path=tmp_path / "old-bang-ke.json",
        sha256="c" * 64,
        status=BatchStatus.READY,
    )
    database.execute(
        """
        UPDATE batches
        SET row_count = 3, valid_count = 3, total_amount = 600000
        WHERE id = ?
        """,
        (batch.id,),
    )
    with database.transaction(immediate=True) as connection:
        run_ids: list[int] = []
        for suffix in (1, 2):
            cursor = connection.execute(
                """
                INSERT INTO excel_runs (
                    operation, started_at, status, total_items,
                    changed_items, skipped_items, conflict_count
                ) VALUES ('EXPENSE_POSTING', ?, 'SUCCEEDED', 3, 2, 1, 0)
                """,
                (f"2026-08-1{suffix}T08:00:00+07:00",),
            )
            run_ids.append(int(cursor.lastrowid))
        for run_id in run_ids:
            rows = (
                (0, "VAT", "VAT", 100_000, "POSTED", "OVERWRITE"),
                (1, " vat ", "VAT", 200_000, "USER_SKIPPED", "SKIP"),
                (2, "VSDL", "VSDL", 300_000, "POSTED", "OVERWRITE"),
            )
            for index, original, selected, amount, status, action in rows:
                connection.execute(
                    """
                    INSERT INTO expense_posting_items (
                        run_id, batch_id, batch_hash, source_item_index,
                        fee_original, fee_selected, rule, amount,
                        action, status, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, 'GV', ?, ?, ?, ?)
                    """,
                    (
                        run_id,
                        batch.id,
                        batch.sha256,
                        index,
                        original,
                        selected,
                        amount,
                        action,
                        status,
                        "2026-08-18T08:00:00+07:00",
                    ),
                )
    database.close()
    connection = sqlite3.connect(path)
    connection.execute("PRAGMA user_version = 16")
    connection.commit()
    connection.close()

    migrated = Database(path)

    assert migrated.query_one("PRAGMA user_version")[0] == SQLITE_SCHEMA_VERSION
    assert migrated.query_one(
        """
        SELECT COUNT(*)
        FROM expense_posting_items
        WHERE UPPER(TRIM(fee_original)) = 'VAT'
           OR UPPER(TRIM(COALESCE(fee_selected, ''))) = 'VAT'
        """
    )[0] == 0
    assert migrated.query_one("SELECT COUNT(*) FROM expense_posting_items")[0] == 2
    updated_batch = migrated.query_one(
        "SELECT row_count, valid_count, total_amount FROM batches WHERE id = ?",
        (batch.id,),
    )
    assert tuple(updated_batch) == (1, 1, 300_000)
    updated_runs = migrated.query_all(
        """
        SELECT total_items, changed_items, skipped_items
        FROM excel_runs
        ORDER BY id
        """
    )
    assert [tuple(row) for row in updated_runs] == [(1, 1, 0), (1, 1, 0)]
    assert migrated.query_one("PRAGMA integrity_check")[0] == "ok"
    assert migrated.query_all("PRAGMA foreign_key_check") == []
    migrated.close()
