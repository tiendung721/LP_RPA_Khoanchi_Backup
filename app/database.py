"""Kết nối SQLite, transaction và migration schema."""

from __future__ import annotations

import json
import sqlite3
import threading
from collections.abc import Iterator, Sequence
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from app.constants import SQLITE_SCHEMA_VERSION


class DatabaseError(RuntimeError):
    """Lỗi tầng lưu trữ đã được chuẩn hóa."""


class Database:
    """Một kết nối SQLite tuần tự hóa bằng khóa re-entrant.

    ``check_same_thread=False`` cho phép watcher/service gọi repository từ thread
    nền; khóa bảo đảm không có hai transaction dùng chung connection đồng thời.
    """

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        self._connection: sqlite3.Connection | None = None
        self.connect()
        self.initialize()

    def connect(self) -> sqlite3.Connection:
        with self._lock:
            if self._connection is None:
                connection = sqlite3.connect(
                    self.path,
                    timeout=30.0,
                    isolation_level=None,
                    check_same_thread=False,
                )
                connection.row_factory = sqlite3.Row
                connection.execute("PRAGMA foreign_keys = ON")
                connection.execute("PRAGMA busy_timeout = 30000")
                try:
                    connection.execute("PRAGMA journal_mode = WAL")
                    connection.execute("PRAGMA synchronous = FULL")
                except sqlite3.DatabaseError:
                    connection.close()
                    raise
                self._connection = connection
            return self._connection

    @property
    def connection(self) -> sqlite3.Connection:
        return self.connect()

    def initialize(self) -> None:
        with self.transaction(immediate=True) as connection:
            current_version = int(
                connection.execute("PRAGMA user_version").fetchone()[0]
            )
            if current_version > SQLITE_SCHEMA_VERSION:
                raise DatabaseError(
                    "Database được tạo bởi phiên bản ứng dụng mới hơn."
                )
            if current_version < 1:
                self._migration_1(connection)
                connection.execute("PRAGMA user_version = 1")
                current_version = 1
            if current_version < 2:
                self._migration_2(connection)
                connection.execute("PRAGMA user_version = 2")
                current_version = 2
            if current_version < 3:
                self._migration_3(connection)
                connection.execute("PRAGMA user_version = 3")
                current_version = 3
            if current_version < 4:
                self._migration_4(connection)
                connection.execute("PRAGMA user_version = 4")
                current_version = 4
            if current_version < 5:
                self._migration_5(connection)
                connection.execute("PRAGMA user_version = 5")
                current_version = 5
            if current_version < 6:
                self._migration_6(connection)
                connection.execute("PRAGMA user_version = 6")
                current_version = 6
            if current_version < 7:
                self._migration_7(connection)
                connection.execute("PRAGMA user_version = 7")
                current_version = 7
            if current_version < 8:
                self._migration_8(connection)
                connection.execute("PRAGMA user_version = 8")
                current_version = 8
            if current_version < 9:
                self._migration_9(connection)
                connection.execute("PRAGMA user_version = 9")
                current_version = 9
            if current_version < 10:
                self._migration_10(connection)
                connection.execute("PRAGMA user_version = 10")
                current_version = 10
            if current_version < 11:
                self._migration_11(connection)
                connection.execute("PRAGMA user_version = 11")
                current_version = 11
            if current_version < 12:
                self._migration_12(connection)
                connection.execute("PRAGMA user_version = 12")
                current_version = 12
            if current_version < 13:
                self._migration_13(connection)
                connection.execute("PRAGMA user_version = 13")
                current_version = 13
            if current_version < 14:
                self._migration_14(connection)
                connection.execute("PRAGMA user_version = 14")
                current_version = 14
            if current_version < 15:
                self._migration_15(connection)
                connection.execute("PRAGMA user_version = 15")
                current_version = 15
            if current_version < 16:
                self._migration_16(connection)
                connection.execute("PRAGMA user_version = 16")
                current_version = 16
            if current_version < 17:
                self._migration_17(connection)
                connection.execute("PRAGMA user_version = 17")
                current_version = 17
            if current_version < 18:
                self._migration_18(connection)
                connection.execute("PRAGMA user_version = 18")
                current_version = 18
            if current_version < 19:
                self._migration_19(connection)
                connection.execute("PRAGMA user_version = 19")
                current_version = 19
            if current_version < 20:
                self._migration_20(connection)
                connection.execute("PRAGMA user_version = 20")
                current_version = 20
            self._ensure_schema_20(connection)
            if current_version != SQLITE_SCHEMA_VERSION:
                raise DatabaseError("Không thể nâng cấp database đến phiên bản hiện tại.")

    migrate = initialize

    @staticmethod
    def _migration_1(connection: sqlite3.Connection) -> None:
        # Dùng từng statement để không kích hoạt implicit COMMIT của
        # sqlite3.Connection.executescript; migration vì thế nằm trọn trong
        # transaction do ``initialize`` quản lý.
        statements = (
            """
            CREATE TABLE IF NOT EXISTS batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_filename TEXT NOT NULL,
                source_inbox_path TEXT,
                original_archive_path TEXT NOT NULL,
                working_path TEXT NOT NULL,
                ready_path TEXT,
                sha256 TEXT NOT NULL UNIQUE,
                status TEXT NOT NULL CHECK (
                    status IN ('RECEIVED','REVIEWING','READY','INVALID','ARCHIVED')
                ),
                received_at TEXT NOT NULL,
                last_opened_at TEXT,
                last_saved_at TEXT,
                confirmed_at TEXT,
                row_count INTEGER NOT NULL DEFAULT 0,
                valid_count INTEGER NOT NULL DEFAULT 0,
                warning_count INTEGER NOT NULL DEFAULT 0,
                error_count INTEGER NOT NULL DEFAULT 0,
                total_amount INTEGER NOT NULL DEFAULT 0,
                last_error TEXT
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS app_state (
                key TEXT PRIMARY KEY,
                value TEXT
            )
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_batches_status_received
                ON batches(status, received_at DESC, id DESC)
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_batches_confirmed
                ON batches(confirmed_at DESC, id DESC)
            """,
        )
        for statement in statements:
            connection.execute(statement)

    @staticmethod
    def _migration_2(connection: sqlite3.Connection) -> None:
        columns = {
            str(row["name"])
            for row in connection.execute("PRAGMA table_info(batches)").fetchall()
        }
        if "source_inbox_path" in columns and "source_output_path" not in columns:
            connection.execute(
                "ALTER TABLE batches RENAME COLUMN source_inbox_path TO source_output_path"
            )
        connection.execute(
            """
            INSERT INTO app_state(key, value)
            SELECT 'last_output_scan_at', value
            FROM app_state
            WHERE key = 'last_inbox_scan_at'
            ON CONFLICT(key) DO NOTHING
            """
        )
        connection.execute("DELETE FROM app_state WHERE key = 'last_inbox_scan_at'")

    @staticmethod
    def _migration_3(connection: sqlite3.Connection) -> None:
        statements = (
            """
            CREATE TABLE IF NOT EXISTS excel_runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                operation TEXT NOT NULL,
                started_at TEXT NOT NULL,
                completed_at TEXT,
                source_path TEXT,
                target_path TEXT,
                source_fingerprint TEXT,
                target_fingerprint_before TEXT,
                target_fingerprint_after TEXT,
                sheet_name TEXT,
                backup_path TEXT,
                status TEXT NOT NULL CHECK (
                    status IN (
                        'ANALYZING', 'WAITING_USER', 'APPLYING', 'SUCCEEDED',
                        'NO_CHANGES', 'CANCELLED', 'FAILED'
                    )
                ),
                total_items INTEGER NOT NULL DEFAULT 0 CHECK (total_items >= 0),
                changed_items INTEGER NOT NULL DEFAULT 0 CHECK (changed_items >= 0),
                skipped_items INTEGER NOT NULL DEFAULT 0 CHECK (skipped_items >= 0),
                conflict_count INTEGER NOT NULL DEFAULT 0 CHECK (conflict_count >= 0),
                error_message TEXT
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS expense_posting_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id INTEGER NOT NULL,
                batch_id INTEGER NOT NULL,
                batch_hash TEXT NOT NULL,
                source_item_index INTEGER NOT NULL CHECK (source_item_index >= 0),
                container TEXT,
                bl TEXT,
                fee_original TEXT NOT NULL,
                fee_selected TEXT,
                rule TEXT,
                amount INTEGER NOT NULL CHECK (amount >= 0),
                sheet_name TEXT,
                target_row INTEGER CHECK (target_row IS NULL OR target_row > 0),
                target_column INTEGER CHECK (
                    target_column IS NULL OR target_column > 0
                ),
                target_cell TEXT,
                value_before TEXT,
                value_after TEXT,
                action TEXT,
                status TEXT NOT NULL CHECK (
                    status IN (
                        'PLANNED', 'POSTED', 'ALREADY_EXISTS', 'USER_SKIPPED',
                        'NOT_MATCHED', 'UNRESOLVED', 'FAILED'
                    )
                ),
                created_at TEXT NOT NULL,
                FOREIGN KEY (run_id) REFERENCES excel_runs(id) ON DELETE CASCADE,
                FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE RESTRICT
            )
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_excel_runs_operation_status
                ON excel_runs(operation, status, started_at DESC, id DESC)
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_expense_posting_items_run
                ON expense_posting_items(run_id, source_item_index)
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_expense_posting_items_batch
                ON expense_posting_items(batch_hash, source_item_index, status)
            """,
            """
            CREATE UNIQUE INDEX IF NOT EXISTS ux_expense_posting_success_source
                ON expense_posting_items(batch_hash, source_item_index)
                WHERE status IN ('POSTED', 'ALREADY_EXISTS')
            """,
        )
        for statement in statements:
            connection.execute(statement)

    @staticmethod
    def _migration_4(connection: sqlite3.Connection) -> None:
        """Cho phép lưu nhiều lần nhập thành công cho cùng một dòng JSON."""

        connection.execute("DROP INDEX IF EXISTS ux_expense_posting_success_source")
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_expense_posting_success_source
                ON expense_posting_items(
                    batch_hash, source_item_index, created_at DESC, id DESC
                )
                WHERE status IN ('POSTED', 'ALREADY_EXISTS')
            """
        )

    @staticmethod
    def _migration_5(connection: sqlite3.Connection) -> None:
        """Phiên bản lịch sử; không tạo dữ liệu tra cứu đã bị loại bỏ."""

        del connection

    @staticmethod
    def _migration_6(connection: sqlite3.Connection) -> None:
        """Loại bỏ bảng job của luồng tra cứu container cũ."""

        connection.execute("DROP TABLE IF EXISTS container_lookup_jobs")

    @staticmethod
    def _migration_7(connection: sqlite3.Connection) -> None:
        """Lưu vết cột và giá trị hóa đơn của luồng nhập khoản chi BK."""

        columns = {
            str(row[1])
            for row in connection.execute(
                "PRAGMA table_info(expense_posting_items)"
            ).fetchall()
        }
        if not columns:
            return
        additions = {
            "invoice_no": "TEXT",
            "invoice_selected": "TEXT",
            "invoice_target_column": "INTEGER",
            "invoice_target_cell": "TEXT",
            "invoice_value_before": "TEXT",
            "invoice_value_after": "TEXT",
            "invoice_action": "TEXT",
        }
        for name, column_type in additions.items():
            if name not in columns:
                connection.execute(
                    f"ALTER TABLE expense_posting_items ADD COLUMN {name} {column_type}"
                )

    @staticmethod
    def _migration_8(connection: sqlite3.Connection) -> None:
        """Ánh xạ dòng kế hoạch được mang từ sheet tháng cũ sang tháng đích."""

        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS expense_carry_forwards (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                workbook_path TEXT NOT NULL,
                source_sheet TEXT NOT NULL,
                source_row INTEGER NOT NULL CHECK (source_row > 0),
                source_sqt INTEGER NOT NULL CHECK (source_sqt > 0),
                container TEXT NOT NULL,
                source_signature TEXT NOT NULL,
                target_sheet TEXT NOT NULL,
                target_row INTEGER NOT NULL CHECK (target_row > 0),
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE (
                    workbook_path, source_sheet, source_row, source_sqt,
                    container, target_sheet
                )
            )
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_expense_carry_forward_target
            ON expense_carry_forwards(workbook_path, target_sheet, target_row)
            """
        )

    @staticmethod
    def _migration_9(connection: sqlite3.Connection) -> None:
        """Lưu vết bên vận tải mà không thay đổi khóa lịch sử hiện có."""

        columns = {
            str(row[1])
            for row in connection.execute(
                "PRAGMA table_info(expense_posting_items)"
            ).fetchall()
        }
        if not columns:
            return
        additions = {
            "carrier_source": "TEXT",
            "carrier_effective": "TEXT",
            "carrier_group": "TEXT",
            "carrier_target_column": "INTEGER",
            "carrier_target_cell": "TEXT",
            "carrier_value_before": "TEXT",
            "carrier_value_after": "TEXT",
            "carrier_action": "TEXT",
        }
        for name, column_type in additions.items():
            if name not in columns:
                connection.execute(
                    f"ALTER TABLE expense_posting_items ADD COLUMN {name} {column_type}"
                )

    @staticmethod
    def _migration_10(connection: sqlite3.Connection) -> None:
        """Nhóm đối soát cước biển theo tàu/chuyến và hóa đơn đóng góp."""

        batch_columns = {
            str(row[1])
            for row in connection.execute("PRAGMA table_info(batches)").fetchall()
        }
        # Một số bản thử nghiệm rất cũ có user_version nhưng thiếu bảng lõi.
        # Tạo lại cấu trúc nền theo cách idempotent trước khi thêm metadata v10.
        if not batch_columns:
            Database._migration_1(connection)
            Database._migration_2(connection)
            batch_columns = {
                str(row[1])
                for row in connection.execute("PRAGMA table_info(batches)").fetchall()
            }
        if "source_kind" not in batch_columns:
            connection.execute(
                "ALTER TABLE batches ADD COLUMN source_kind TEXT NOT NULL "
                "DEFAULT 'ASSISTANT'"
            )
        if "reconciliation_group_id" not in batch_columns:
            connection.execute(
                "ALTER TABLE batches ADD COLUMN reconciliation_group_id INTEGER"
            )

        statements = (
            """
            CREATE TABLE IF NOT EXISTS sea_freight_reconciliation_groups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bk_path TEXT NOT NULL,
                bk_sheet TEXT NOT NULL,
                vessel_voyage_raw TEXT NOT NULL,
                vessel_key TEXT NOT NULL,
                voyage_key TEXT NOT NULL,
                combined_key TEXT NOT NULL,
                bk_container_count INTEGER NOT NULL DEFAULT 0 CHECK (bk_container_count >= 0),
                received_container_count INTEGER NOT NULL DEFAULT 0 CHECK (received_container_count >= 0),
                total_amount INTEGER NOT NULL DEFAULT 0 CHECK (total_amount >= 0),
                status TEXT NOT NULL CHECK (status IN (
                    'PENDING','READY','OVER_COUNT','BK_NOT_FOUND',
                    'METADATA_CONFLICT','NEEDS_RECHECK','ALLOCATED',
                    'POSTED','CANCELLED'
                )),
                bk_fingerprint TEXT,
                container_snapshot_hash TEXT,
                output_carrier TEXT,
                invoice_conflict_accepted INTEGER NOT NULL DEFAULT 0,
                generated_batch_id INTEGER,
                posted_run_id INTEGER,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                posted_at TEXT,
                FOREIGN KEY (generated_batch_id) REFERENCES batches(id) ON DELETE SET NULL,
                FOREIGN KEY (posted_run_id) REFERENCES excel_runs(id) ON DELETE SET NULL
            )
            """,
            """
            CREATE UNIQUE INDEX IF NOT EXISTS ux_sea_freight_open_group
            ON sea_freight_reconciliation_groups(
                bk_path, bk_sheet, vessel_key, voyage_key
            )
            WHERE status NOT IN ('POSTED','CANCELLED')
            """,
            """
            CREATE TABLE IF NOT EXISTS sea_freight_invoice_contributions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL,
                source_batch_id INTEGER,
                source_item_index INTEGER NOT NULL CHECK (source_item_index >= 0),
                source_sha256 TEXT NOT NULL,
                invoice_no TEXT,
                invoice_date TEXT,
                bl TEXT,
                vessel_voyage_raw TEXT NOT NULL,
                vessel_name TEXT NOT NULL,
                voyage_no TEXT NOT NULL,
                invoice_container_count INTEGER NOT NULL CHECK (invoice_container_count > 0),
                container_count_basis TEXT NOT NULL CHECK (
                    container_count_basis IN ('EXPLICIT','CALCULATED')
                ),
                carrier TEXT,
                amount INTEGER NOT NULL CHECK (amount >= 0),
                fingerprint TEXT NOT NULL,
                status TEXT NOT NULL CHECK (status IN ('ACTIVE','REMOVED','DUPLICATE')),
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                removed_at TEXT,
                FOREIGN KEY (group_id) REFERENCES sea_freight_reconciliation_groups(id) ON DELETE CASCADE,
                FOREIGN KEY (source_batch_id) REFERENCES batches(id) ON DELETE SET NULL
            )
            """,
            """
            CREATE UNIQUE INDEX IF NOT EXISTS ux_sea_freight_contribution_fingerprint
            ON sea_freight_invoice_contributions(fingerprint)
            WHERE status != 'REMOVED'
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_sea_freight_contribution_group
            ON sea_freight_invoice_contributions(group_id, status, id)
            """,
            """
            CREATE TABLE IF NOT EXISTS sea_freight_group_containers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL,
                container TEXT NOT NULL,
                source_sheet TEXT NOT NULL,
                source_row INTEGER NOT NULL CHECK (source_row > 0),
                source_sqt INTEGER,
                departure_date TEXT,
                allocation_order INTEGER NOT NULL CHECK (allocation_order >= 0),
                FOREIGN KEY (group_id) REFERENCES sea_freight_reconciliation_groups(id) ON DELETE CASCADE,
                UNIQUE(group_id, container)
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS sea_freight_reconciliation_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL,
                event_type TEXT NOT NULL,
                payload_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (group_id) REFERENCES sea_freight_reconciliation_groups(id) ON DELETE CASCADE
            )
            """,
            """
            CREATE INDEX IF NOT EXISTS idx_sea_freight_group_status
            ON sea_freight_reconciliation_groups(status, updated_at DESC, id DESC)
            """,
        )
        for statement in statements:
            connection.execute(statement)

    @staticmethod
    def _ensure_schema_10(connection: sqlite3.Connection) -> None:
        """Bổ sung cột v10 nhỏ theo cách idempotent cho các bản beta đã migrate."""

        columns = {
            str(row[1])
            for row in connection.execute(
                "PRAGMA table_info(sea_freight_reconciliation_groups)"
            ).fetchall()
        }
        if columns and "invoice_conflict_accepted" not in columns:
            connection.execute(
                "ALTER TABLE sea_freight_reconciliation_groups "
                "ADD COLUMN invoice_conflict_accepted INTEGER NOT NULL DEFAULT 0"
            )

    @staticmethod
    def _migration_11(connection: sqlite3.Connection) -> None:
        """Bổ sung kỳ đối soát và nguồn HĐ cho cửa sổ đối soát thống nhất."""

        Database._ensure_schema_10(connection)
        group_columns = {
            str(row[1])
            for row in connection.execute(
                "PRAGMA table_info(sea_freight_reconciliation_groups)"
            ).fetchall()
        }
        for name, definition in (
            ("reconciliation_month", "INTEGER CHECK (reconciliation_month BETWEEN 1 AND 12)"),
            ("reconciliation_year", "INTEGER CHECK (reconciliation_year BETWEEN 2000 AND 2999)"),
            ("bk_issue", "TEXT"),
            ("bk_invalid_container_count", "INTEGER NOT NULL DEFAULT 0"),
            ("bk_duplicate_container_count", "INTEGER NOT NULL DEFAULT 0"),
        ):
            if name not in group_columns:
                connection.execute(
                    f"ALTER TABLE sea_freight_reconciliation_groups ADD COLUMN {name} {definition}"
                )
        contribution_columns = {
            str(row[1])
            for row in connection.execute(
                "PRAGMA table_info(sea_freight_invoice_contributions)"
            ).fetchall()
        }
        if "source_kind" not in contribution_columns:
            connection.execute(
                "ALTER TABLE sea_freight_invoice_contributions "
                "ADD COLUMN source_kind TEXT NOT NULL DEFAULT 'INITIAL'"
            )
        connection.execute(
            """
            UPDATE sea_freight_reconciliation_groups
            SET reconciliation_month = CAST(SUBSTR(TRIM(bk_sheet), 2, 2) AS INTEGER),
                reconciliation_year = 2000 + CAST(SUBSTR(TRIM(bk_sheet), 5, 2) AS INTEGER)
            WHERE reconciliation_month IS NULL
              AND TRIM(bk_sheet) GLOB 'T[0-1][0-9] [0-9][0-9]'
            """
        )
        connection.execute(
            "CREATE INDEX IF NOT EXISTS idx_sea_freight_group_period "
            "ON sea_freight_reconciliation_groups(reconciliation_year, reconciliation_month, updated_at DESC)"
        )

    @staticmethod
    def _ensure_schema_11(connection: sqlite3.Connection) -> None:
        """Tự sửa các database beta v11 bị tạo thiếu cột."""

        Database._migration_11(connection)

    @staticmethod
    def _migration_12(connection: sqlite3.Connection) -> None:
        """Giải phóng HĐ của hồ sơ đã hủy nhưng vẫn giữ nguyên lịch sử."""

        Database._ensure_schema_11(connection)
        connection.execute(
            """
            UPDATE sea_freight_invoice_contributions
            SET status = 'REMOVED',
                removed_at = COALESCE(
                    removed_at,
                    (SELECT g.updated_at
                     FROM sea_freight_reconciliation_groups AS g
                     WHERE g.id = sea_freight_invoice_contributions.group_id)
                ),
                updated_at = COALESCE(
                    (SELECT g.updated_at
                     FROM sea_freight_reconciliation_groups AS g
                     WHERE g.id = sea_freight_invoice_contributions.group_id),
                    updated_at
                )
            WHERE status = 'ACTIVE'
              AND group_id IN (
                  SELECT id FROM sea_freight_reconciliation_groups
                  WHERE status = 'CANCELLED'
              )
            """
        )
        batch_columns = {
            str(row[1])
            for row in connection.execute("PRAGMA table_info(batches)").fetchall()
        }
        if {"status", "reconciliation_group_id"}.issubset(batch_columns):
            connection.execute(
                """
                UPDATE batches
                SET status = 'ARCHIVED'
                WHERE reconciliation_group_id IN (
                    SELECT id FROM sea_freight_reconciliation_groups
                    WHERE status = 'CANCELLED'
                )
                  AND status IN ('RECEIVED','REVIEWING','READY')
                """
            )

    @staticmethod
    def _ensure_schema_12(connection: sqlite3.Connection) -> None:
        Database._migration_12(connection)

    @staticmethod
    def _migration_13(connection: sqlite3.Connection) -> None:
        """Phiên đối soát bất biến và liên kết rõ với dòng file nguồn."""

        Database._ensure_schema_12(connection)
        group_columns = {
            str(row[1])
            for row in connection.execute(
                "PRAGMA table_info(sea_freight_reconciliation_groups)"
            ).fetchall()
        }
        additions = (
            ("revision_no", "INTEGER NOT NULL DEFAULT 1 CHECK (revision_no > 0)"),
            ("supersedes_group_id", "INTEGER"),
            ("is_current", "INTEGER NOT NULL DEFAULT 1 CHECK (is_current IN (0,1))"),
            ("primary_source_batch_id", "INTEGER"),
            ("primary_source_item_index", "INTEGER"),
        )
        for name, definition in additions:
            if name not in group_columns:
                connection.execute(
                    f"ALTER TABLE sea_freight_reconciliation_groups ADD COLUMN {name} {definition}"
                )

        revision_rows = connection.execute(
            """
            SELECT id, bk_path, bk_sheet, vessel_key, voyage_key
            FROM sea_freight_reconciliation_groups
            ORDER BY bk_path, bk_sheet, vessel_key, voyage_key, id
            """
        ).fetchall()
        previous_key: tuple[str, str, str, str] | None = None
        previous_id: int | None = None
        revision_no = 0
        for row in revision_rows:
            key = (str(row[1]), str(row[2]), str(row[3]), str(row[4]))
            if key != previous_key:
                previous_key = key
                previous_id = None
                revision_no = 1
            else:
                revision_no += 1
            connection.execute(
                """
                UPDATE sea_freight_reconciliation_groups
                SET revision_no = ?, supersedes_group_id = ?
                WHERE id = ?
                """,
                (revision_no, previous_id, int(row[0])),
            )
            previous_id = int(row[0])

        # Hồ sơ mới nhất chưa hủy của cùng tàu/chuyến là phiên hiện hành.
        connection.execute(
            "UPDATE sea_freight_reconciliation_groups SET is_current = 0"
        )
        connection.execute(
            """
            UPDATE sea_freight_reconciliation_groups AS current
            SET is_current = 1
            WHERE current.id = (
                SELECT candidate.id
                FROM sea_freight_reconciliation_groups AS candidate
                WHERE candidate.bk_path = current.bk_path
                  AND candidate.bk_sheet = current.bk_sheet
                  AND candidate.vessel_key = current.vessel_key
                  AND candidate.voyage_key = current.voyage_key
                  AND candidate.status != 'CANCELLED'
                ORDER BY candidate.id DESC
                LIMIT 1
            )
            """
        )
        connection.execute(
            """
            UPDATE sea_freight_reconciliation_groups
            SET primary_source_batch_id = (
                    SELECT c.source_batch_id
                    FROM sea_freight_invoice_contributions AS c
                    WHERE c.group_id = sea_freight_reconciliation_groups.id
                      AND c.source_batch_id IS NOT NULL
                    ORDER BY CASE c.source_kind WHEN 'INITIAL' THEN 0 ELSE 1 END, c.id
                    LIMIT 1
                ),
                primary_source_item_index = (
                    SELECT c.source_item_index
                    FROM sea_freight_invoice_contributions AS c
                    WHERE c.group_id = sea_freight_reconciliation_groups.id
                      AND c.source_batch_id IS NOT NULL
                    ORDER BY CASE c.source_kind WHEN 'INITIAL' THEN 0 ELSE 1 END, c.id
                    LIMIT 1
                )
            WHERE primary_source_batch_id IS NULL
            """
        )

        # Phiên mới được phép giữ cùng fingerprint HĐ; chống trùng trong từng hồ sơ.
        connection.execute("DROP INDEX IF EXISTS ux_sea_freight_contribution_fingerprint")
        connection.execute(
            """
            CREATE UNIQUE INDEX IF NOT EXISTS ux_sea_freight_group_contribution_fingerprint
            ON sea_freight_invoice_contributions(group_id, fingerprint)
            WHERE status != 'REMOVED'
            """
        )
        connection.execute("DROP INDEX IF EXISTS ux_sea_freight_open_group")
        connection.execute(
            """
            CREATE UNIQUE INDEX IF NOT EXISTS ux_sea_freight_current_open_group
            ON sea_freight_reconciliation_groups(
                bk_path, bk_sheet, vessel_key, voyage_key
            )
            WHERE is_current = 1 AND status != 'CANCELLED'
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_sea_freight_source_owner
            ON sea_freight_reconciliation_groups(
                primary_source_batch_id, is_current, status, id
            )
            """
        )
        posting_columns = {
            str(row[1])
            for row in connection.execute(
                "PRAGMA table_info(expense_posting_items)"
            ).fetchall()
        }
        if {"batch_id", "source_item_index", "status", "id"}.issubset(posting_columns):
            connection.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_expense_posting_origin
                ON expense_posting_items(batch_id, source_item_index, status, id)
                """
            )

        # Batch con không còn là file làm việc hiện hành và phiên cũ không được
        # tự xuất hiện trong danh sách READY để ghi nhầm.
        batch_columns = {
            str(row[1])
            for row in connection.execute("PRAGMA table_info(batches)").fetchall()
        }
        if {"status", "reconciliation_group_id"}.issubset(batch_columns):
            connection.execute(
                """
                UPDATE batches
                SET status = 'ARCHIVED'
                WHERE reconciliation_group_id IN (
                    SELECT id FROM sea_freight_reconciliation_groups
                    WHERE is_current = 0 OR status = 'CANCELLED'
                )
                  AND status IN ('RECEIVED','REVIEWING','READY','INVALID')
                """
            )
            active_row = connection.execute(
                "SELECT value FROM app_state WHERE key = 'active_batch_id'"
            ).fetchone()
            if active_row is not None and active_row[0] not in (None, ""):
                try:
                    active_batch_id = int(active_row[0])
                except (TypeError, ValueError):
                    active_batch_id = 0
                owner = connection.execute(
                    """
                    SELECT g.primary_source_batch_id
                    FROM batches AS b
                    JOIN sea_freight_reconciliation_groups AS g
                      ON g.id = b.reconciliation_group_id
                    WHERE b.id = ?
                    """,
                    (active_batch_id,),
                ).fetchone()
                if owner is not None and owner[0] is not None:
                    connection.execute(
                        "UPDATE app_state SET value = ? WHERE key = 'active_batch_id'",
                        (str(int(owner[0])),),
                    )

    @staticmethod
    def _ensure_schema_13(connection: sqlite3.Connection) -> None:
        Database._migration_13(connection)

    @staticmethod
    def _migration_14(connection: sqlite3.Connection) -> None:
        """Lưu lựa chọn của phiên Excel chưa hoàn tất để có thể tiếp tục."""

        Database._ensure_schema_13(connection)
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS excel_resolution_drafts (
                context_key TEXT PRIMARY KEY,
                operation TEXT NOT NULL,
                context_json TEXT NOT NULL,
                payload_json TEXT NOT NULL,
                status TEXT NOT NULL CHECK (
                    status IN ('ACTIVE', 'FAILED', 'COMPLETED')
                ),
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                completed_at TEXT,
                last_error TEXT
            )
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_excel_resolution_drafts_status
            ON excel_resolution_drafts(operation, status, updated_at DESC)
            """
        )

    @staticmethod
    def _ensure_schema_14(connection: sqlite3.Connection) -> None:
        Database._migration_14(connection)

    @staticmethod
    def _migration_15(connection: sqlite3.Connection) -> None:
        """Giữ lần xử lý xung đột gần nhất theo file nguồn và nghiệp vụ."""

        Database._ensure_schema_14(connection)
        Database._ensure_schema_15(connection)

        rows = connection.execute(
            "SELECT * FROM excel_resolution_drafts ORDER BY updated_at, context_key"
        ).fetchall()
        for row in rows:
            try:
                context = json.loads(str(row["context_json"]))
            except (TypeError, ValueError, json.JSONDecodeError):
                context = {}
            operation = str(row["operation"])
            batch_id = context.get("batch_id")
            source_path = context.get("source_path")
            source_identity = context.get("source_identity")
            if batch_id not in (None, ""):
                source_file_key = f"{operation}:batch:{batch_id}"
            elif source_path not in (None, ""):
                source_file_key = f"{operation}:path:{str(source_path).casefold()}"
            elif source_identity not in (None, ""):
                source_file_key = f"{operation}:identity:{source_identity}"
            else:
                source_file_key = f"{operation}:legacy:{row['context_key']}"
            status = (
                "SUCCEEDED" if str(row["status"]) == "COMPLETED" else str(row["status"])
            )
            connection.execute(
                """
                INSERT INTO excel_resolution_latest (
                    source_file_key, operation, context_json, payload_json,
                    status, run_id, created_at, updated_at, completed_at,
                    last_error
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(source_file_key) DO UPDATE SET
                    operation = excluded.operation,
                    context_json = excluded.context_json,
                    payload_json = excluded.payload_json,
                    status = excluded.status,
                    run_id = excluded.run_id,
                    created_at = excluded.created_at,
                    updated_at = excluded.updated_at,
                    completed_at = excluded.completed_at,
                    last_error = excluded.last_error
                WHERE excluded.updated_at >= excel_resolution_latest.updated_at
                """,
                (
                    source_file_key,
                    operation,
                    str(row["context_json"]),
                    str(row["payload_json"]),
                    status,
                    context.get("run_id"),
                    str(row["created_at"]),
                    str(row["updated_at"]),
                    row["completed_at"],
                    row["last_error"],
                ),
            )

    @staticmethod
    def _ensure_schema_15(connection: sqlite3.Connection) -> None:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS excel_resolution_latest (
                source_file_key TEXT PRIMARY KEY,
                operation TEXT NOT NULL,
                context_json TEXT NOT NULL,
                payload_json TEXT NOT NULL,
                status TEXT NOT NULL CHECK (
                    status IN ('ACTIVE', 'SUCCEEDED', 'FAILED', 'CANCELLED')
                ),
                run_id INTEGER,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                completed_at TEXT,
                last_error TEXT
            )
            """
        )

    @staticmethod
    def _migration_16(connection: sqlite3.Connection) -> None:
        """Lưu nguồn chứng từ gốc cho từng hóa đơn đối soát cước biển."""

        Database._ensure_schema_15(connection)
        Database._ensure_schema_16(connection)
        batch_columns = {
            str(row[1])
            for row in connection.execute("PRAGMA table_info(batches)").fetchall()
        }
        source_name_expression = (
            "COALESCE(NULLIF((SELECT b.source_filename FROM batches AS b "
            "WHERE b.id = sea_freight_invoice_contributions.source_batch_id), ''), "
            "'Dữ liệu bóc tách cũ')"
            if "source_filename" in batch_columns
            else "'Dữ liệu bóc tách cũ'"
        )
        connection.execute(
            f"""
            UPDATE sea_freight_invoice_contributions
            SET source_document_id = CASE
                    WHEN source_document_id != 'LEGACY_DOCUMENT'
                        THEN source_document_id
                    WHEN source_batch_id IS NOT NULL
                        THEN 'LEGACY_BATCH_' || CAST(source_batch_id AS TEXT)
                    ELSE 'LEGACY_DOCUMENT'
                END,
                source_document_name = CASE
                    WHEN source_document_name != 'Dữ liệu bóc tách cũ'
                        THEN source_document_name
                    ELSE {source_name_expression}
                END
            WHERE source_document_id = 'LEGACY_DOCUMENT'
               OR source_document_name = 'Dữ liệu bóc tách cũ'
            """
        )

    @staticmethod
    def _ensure_schema_16(connection: sqlite3.Connection) -> None:
        Database._ensure_schema_15(connection)
        columns = {
            str(row[1])
            for row in connection.execute(
                "PRAGMA table_info(sea_freight_invoice_contributions)"
            ).fetchall()
        }
        if columns and "source_document_id" not in columns:
            connection.execute(
                "ALTER TABLE sea_freight_invoice_contributions "
                "ADD COLUMN source_document_id TEXT NOT NULL DEFAULT 'LEGACY_DOCUMENT'"
            )
        if columns and "source_document_name" not in columns:
            connection.execute(
                "ALTER TABLE sea_freight_invoice_contributions "
                "ADD COLUMN source_document_name TEXT NOT NULL DEFAULT 'Dữ liệu bóc tách cũ'"
            )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_excel_resolution_latest_operation
            ON excel_resolution_latest(operation, updated_at DESC)
            """
        )

    @staticmethod
    def _migration_17(connection: sqlite3.Connection) -> None:
        """Loại phí VAT khỏi lịch sử nhập khoản chi và sửa số liệu tổng hợp."""

        Database._ensure_schema_16(connection)
        posting_table = connection.execute(
            """
            SELECT 1
            FROM sqlite_master
            WHERE type = 'table' AND name = 'expense_posting_items'
            """
        ).fetchone()
        if posting_table is None:
            return

        vat_condition = """
            UPPER(TRIM(fee_original)) = 'VAT'
            OR UPPER(TRIM(COALESCE(fee_selected, ''))) = 'VAT'
        """
        connection.execute("DROP TABLE IF EXISTS temp.vat_cleanup_runs")
        connection.execute("DROP TABLE IF EXISTS temp.vat_cleanup_batches")
        connection.execute(
            f"""
            CREATE TEMP TABLE vat_cleanup_runs AS
            SELECT DISTINCT run_id
            FROM expense_posting_items
            WHERE {vat_condition}
            """
        )
        connection.execute(
            f"""
            CREATE TEMP TABLE vat_cleanup_batches AS
            SELECT batch_id,
                   COUNT(*) AS removed_rows,
                   COALESCE(SUM(amount), 0) AS removed_amount
            FROM (
                SELECT batch_id, source_item_index, MAX(amount) AS amount
                FROM expense_posting_items
                WHERE {vat_condition}
                GROUP BY batch_id, source_item_index
            )
            GROUP BY batch_id
            """
        )
        batch_columns = {
            str(row[1])
            for row in connection.execute("PRAGMA table_info(batches)").fetchall()
        }
        if {"id", "row_count", "valid_count", "total_amount"}.issubset(
            batch_columns
        ):
            connection.execute(
                """
                UPDATE batches
                SET row_count = MAX(
                        0,
                        row_count - COALESCE(
                            (SELECT removed_rows FROM vat_cleanup_batches
                             WHERE batch_id = batches.id),
                            0
                        )
                    ),
                    valid_count = MAX(
                        0,
                        valid_count - COALESCE(
                            (SELECT removed_rows FROM vat_cleanup_batches
                             WHERE batch_id = batches.id),
                            0
                        )
                    ),
                    total_amount = MAX(
                        0,
                        total_amount - COALESCE(
                            (SELECT removed_amount FROM vat_cleanup_batches
                             WHERE batch_id = batches.id),
                            0
                        )
                    )
                WHERE id IN (SELECT batch_id FROM vat_cleanup_batches)
                """
            )
        connection.execute(
            f"DELETE FROM expense_posting_items WHERE {vat_condition}"
        )
        connection.execute(
            """
            UPDATE excel_runs
            SET total_items = (
                    SELECT COUNT(*)
                    FROM expense_posting_items AS item
                    WHERE item.run_id = excel_runs.id
                ),
                changed_items = (
                    SELECT COUNT(*)
                    FROM expense_posting_items AS item
                    WHERE item.run_id = excel_runs.id
                      AND item.status = 'POSTED'
                ),
                skipped_items = (
                    SELECT COUNT(*)
                    FROM expense_posting_items AS item
                    WHERE item.run_id = excel_runs.id
                      AND item.status IN (
                          'USER_SKIPPED', 'NOT_MATCHED', 'UNRESOLVED', 'FAILED'
                      )
                )
            WHERE id IN (SELECT run_id FROM vat_cleanup_runs)
            """
        )
        connection.execute("DROP TABLE temp.vat_cleanup_runs")
        connection.execute("DROP TABLE temp.vat_cleanup_batches")

    @staticmethod
    def _ensure_schema_17(connection: sqlite3.Connection) -> None:
        Database._ensure_schema_16(connection)

    @staticmethod
    def _migration_18(connection: sqlite3.Connection) -> None:
        """Lưu manifest chi tiết để có thể xem lại kết quả Excel gần nhất."""

        Database._ensure_schema_17(connection)
        Database._ensure_schema_18(connection)

    @staticmethod
    def _ensure_schema_18(connection: sqlite3.Connection) -> None:
        Database._ensure_schema_17(connection)
        columns = {
            str(row[1])
            for row in connection.execute("PRAGMA table_info(excel_runs)").fetchall()
        }
        if columns and "item_outcomes" not in columns:
            connection.execute("ALTER TABLE excel_runs ADD COLUMN item_outcomes TEXT")

    @staticmethod
    def _migration_19(connection: sqlite3.Connection) -> None:
        """Tăng tốc tra cứu lịch sử HĐ cước biển xuyên batch."""

        Database._ensure_schema_18(connection)
        Database._ensure_schema_19(connection)

    @staticmethod
    def _ensure_schema_19(connection: sqlite3.Connection) -> None:
        Database._ensure_schema_18(connection)
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_sea_freight_contribution_invoice
            ON sea_freight_invoice_contributions(
                UPPER(TRIM(invoice_no)), status, group_id
            )
            """
        )

    @staticmethod
    def _migration_20(connection: sqlite3.Connection) -> None:
        """Nguồn BK nhiều tháng và mọi vị trí xuất hiện của container."""

        Database._ensure_schema_19(connection)
        Database._ensure_schema_20(connection)
        connection.execute(
            """
            INSERT OR IGNORE INTO sea_freight_reconciliation_sources(
                group_id, bk_sheet, reconciliation_month, reconciliation_year,
                vessel_voyage_raw, vessel_name, voyage_no,
                vessel_key, voyage_key, combined_key,
                resolution_kind, container_count, invalid_container_count,
                duplicate_container_count, snapshot_hash, source_order
            )
            SELECT id, bk_sheet, reconciliation_month, reconciliation_year,
                   vessel_voyage_raw, vessel_key, voyage_key,
                   vessel_key, voyage_key, combined_key,
                   CASE WHEN bk_container_count > 0 THEN 'EXACT' ELSE 'NOT_FOUND' END,
                   bk_container_count, bk_invalid_container_count,
                   bk_duplicate_container_count, container_snapshot_hash, 0
            FROM sea_freight_reconciliation_groups
            """
        )
        connection.execute(
            """
            INSERT OR IGNORE INTO sea_freight_container_occurrences(
                group_id, source_id, container, source_sheet, source_row,
                source_sqt, departure_date, selected
            )
            SELECT c.group_id, s.id, c.container, c.source_sheet, c.source_row,
                   c.source_sqt, c.departure_date, 1
            FROM sea_freight_group_containers AS c
            JOIN sea_freight_reconciliation_sources AS s
              ON s.group_id = c.group_id AND s.bk_sheet = c.source_sheet
            """
        )

    @staticmethod
    def _ensure_schema_20(connection: sqlite3.Connection) -> None:
        Database._ensure_schema_19(connection)
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS sea_freight_reconciliation_sources (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL,
                bk_sheet TEXT NOT NULL,
                reconciliation_month INTEGER,
                reconciliation_year INTEGER,
                vessel_voyage_raw TEXT NOT NULL,
                vessel_name TEXT NOT NULL,
                voyage_no TEXT NOT NULL,
                vessel_key TEXT NOT NULL,
                voyage_key TEXT NOT NULL,
                combined_key TEXT NOT NULL,
                resolution_kind TEXT NOT NULL CHECK (resolution_kind IN (
                    'EXACT','AUTO_ALIAS','AMBIGUOUS','NOT_FOUND'
                )),
                container_count INTEGER NOT NULL DEFAULT 0,
                invalid_container_count INTEGER NOT NULL DEFAULT 0,
                duplicate_container_count INTEGER NOT NULL DEFAULT 0,
                snapshot_hash TEXT,
                source_order INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY (group_id) REFERENCES sea_freight_reconciliation_groups(id) ON DELETE CASCADE,
                UNIQUE(group_id, bk_sheet)
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS sea_freight_container_occurrences (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL,
                source_id INTEGER NOT NULL,
                container TEXT NOT NULL,
                source_sheet TEXT NOT NULL,
                source_row INTEGER NOT NULL,
                source_sqt INTEGER,
                departure_date TEXT,
                selected INTEGER NOT NULL DEFAULT 0 CHECK (selected IN (0,1)),
                manually_selected INTEGER NOT NULL DEFAULT 0 CHECK (manually_selected IN (0,1)),
                FOREIGN KEY (group_id) REFERENCES sea_freight_reconciliation_groups(id) ON DELETE CASCADE,
                FOREIGN KEY (source_id) REFERENCES sea_freight_reconciliation_sources(id) ON DELETE CASCADE,
                UNIQUE(group_id, container, source_sheet, source_row)
            )
            """
        )
        connection.execute(
            """
            CREATE UNIQUE INDEX IF NOT EXISTS ux_sea_freight_selected_occurrence
            ON sea_freight_container_occurrences(group_id, container)
            WHERE selected = 1
            """
        )
        connection.execute(
            "CREATE INDEX IF NOT EXISTS idx_sea_freight_source_group "
            "ON sea_freight_reconciliation_sources(group_id, source_order, id)"
        )
        occurrence_columns = {
            str(row[1])
            for row in connection.execute(
                "PRAGMA table_info(sea_freight_container_occurrences)"
            ).fetchall()
        }
        if "manually_selected" not in occurrence_columns:
            connection.execute(
                "ALTER TABLE sea_freight_container_occurrences "
                "ADD COLUMN manually_selected INTEGER NOT NULL DEFAULT 0"
            )

    @contextmanager
    def transaction(
        self,
        *,
        immediate: bool = False,
    ) -> Iterator[sqlite3.Connection]:
        with self._lock:
            connection = self.connect()
            try:
                connection.execute("BEGIN IMMEDIATE" if immediate else "BEGIN")
                yield connection
            except BaseException:
                connection.rollback()
                raise
            else:
                connection.commit()

    def query_one(
        self,
        sql: str,
        parameters: Sequence[Any] = (),
    ) -> sqlite3.Row | None:
        with self._lock:
            return self.connect().execute(sql, tuple(parameters)).fetchone()

    def query_all(
        self,
        sql: str,
        parameters: Sequence[Any] = (),
    ) -> list[sqlite3.Row]:
        with self._lock:
            return list(self.connect().execute(sql, tuple(parameters)).fetchall())

    def execute(
        self,
        sql: str,
        parameters: Sequence[Any] = (),
    ) -> int:
        with self.transaction(immediate=True) as connection:
            cursor = connection.execute(sql, tuple(parameters))
            return cursor.rowcount

    def rebase_paths(self, old_root: str | Path, new_root: str | Path) -> int:
        """Đổi các đường dẫn batch nằm dưới gốc cũ sang vị trí bundle mới."""

        old = Path(old_root).expanduser().resolve(strict=False)
        new = Path(new_root).expanduser().resolve(strict=False)
        if old == new:
            return 0
        changed_rows = 0
        path_columns = (
            "source_output_path",
            "original_archive_path",
            "working_path",
            "ready_path",
        )
        with self.transaction(immediate=True) as connection:
            rows = connection.execute(
                f"SELECT id, {', '.join(path_columns)} FROM batches"
            ).fetchall()
            for row in rows:
                values = {
                    column: _rebased_path_text(row[column], old, new)
                    for column in path_columns
                }
                if all(values[column] == row[column] for column in path_columns):
                    continue
                connection.execute(
                    """
                    UPDATE batches
                    SET source_output_path = ?,
                        original_archive_path = ?,
                        working_path = ?,
                        ready_path = ?
                    WHERE id = ?
                    """,
                    (*(values[column] for column in path_columns), int(row["id"])),
                )
                changed_rows += 1
            carry_rows = connection.execute(
                "SELECT id, workbook_path FROM expense_carry_forwards"
            ).fetchall()
            for row in carry_rows:
                rebased = _rebased_path_text(row["workbook_path"], old, new)
                if rebased == row["workbook_path"]:
                    continue
                connection.execute(
                    "UPDATE expense_carry_forwards SET workbook_path = ? WHERE id = ?",
                    (rebased, int(row["id"])),
                )
                changed_rows += 1
            if connection.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name='sea_freight_reconciliation_groups'"
            ).fetchone():
                group_rows = connection.execute(
                    "SELECT id, bk_path FROM sea_freight_reconciliation_groups"
                ).fetchall()
                for row in group_rows:
                    rebased = _rebased_path_text(row["bk_path"], old, new)
                    if rebased == row["bk_path"]:
                        continue
                    connection.execute(
                        "UPDATE sea_freight_reconciliation_groups SET bk_path = ? WHERE id = ?",
                        (rebased, int(row["id"])),
                    )
                    changed_rows += 1
        return changed_rows

    def close(self) -> None:
        with self._lock:
            if self._connection is not None:
                self._connection.close()
                self._connection = None

    def __enter__(self) -> "Database":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()


def _rebased_path_text(
    value: object,
    old_root: Path,
    new_root: Path,
) -> str | None:
    if value is None:
        return None
    text = str(value)
    try:
        relative = Path(text).resolve(strict=False).relative_to(old_root)
    except ValueError:
        return text
    return str(new_root / relative)
