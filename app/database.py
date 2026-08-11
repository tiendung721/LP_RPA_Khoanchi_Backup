"""Kết nối SQLite, transaction và migration schema."""

from __future__ import annotations

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
