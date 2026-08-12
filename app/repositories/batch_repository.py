"""Repository SQLite cho lịch sử batch và trạng thái ứng dụng."""

from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

from app.constants import APP_STATE_ACTIVE_BATCH_ID
from app.database import Database
from app.models import BatchMetadata, BatchStatus, ValidationSummary

_UPDATABLE_COLUMNS = frozenset(
    {
        "source_filename",
        "source_output_path",
        "original_archive_path",
        "working_path",
        "ready_path",
        "status",
        "last_opened_at",
        "last_saved_at",
        "confirmed_at",
        "row_count",
        "valid_count",
        "warning_count",
        "error_count",
        "total_amount",
        "last_error",
        "received_at",
        "sha256",
        "source_kind",
        "reconciliation_group_id",
    }
)


def local_now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _path_text(value: str | Path | None) -> str | None:
    return None if value is None else str(value)


class BatchRepository:
    def __init__(self, database: Database | str | Path) -> None:
        self.database = (
            database if isinstance(database, Database) else Database(database)
        )

    def create_batch(
        self,
        *,
        source_filename: str,
        source_output_path: str | Path | None,
        original_archive_path: str | Path,
        working_path: str | Path,
        sha256: str,
        status: BatchStatus | str = BatchStatus.RECEIVED,
        received_at: str | None = None,
        ready_path: str | Path | None = None,
        last_error: str | None = None,
        source_kind: str = "ASSISTANT",
        reconciliation_group_id: int | None = None,
    ) -> BatchMetadata:
        status_value = BatchStatus(status).value
        timestamp = received_at or local_now_iso()
        with self.database.transaction(immediate=True) as connection:
            cursor = connection.execute(
                """
                INSERT INTO batches (
                    source_filename, source_output_path, original_archive_path,
                    working_path, ready_path, sha256, status, received_at,
                    last_error, source_kind, reconciliation_group_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    source_filename,
                    _path_text(source_output_path),
                    _path_text(original_archive_path),
                    _path_text(working_path),
                    _path_text(ready_path),
                    sha256,
                    status_value,
                    timestamp,
                    last_error,
                    source_kind,
                    reconciliation_group_id,
                ),
            )
            batch_id = int(cursor.lastrowid)
            row = connection.execute(
                "SELECT * FROM batches WHERE id = ?", (batch_id,)
            ).fetchone()
        if row is None:
            raise RuntimeError("Không thể đọc lại batch vừa tạo.")
        return self._to_metadata(row)

    reserve_batch = create_batch

    def get_by_id(self, batch_id: int) -> BatchMetadata | None:
        row = self.database.query_one(
            "SELECT * FROM batches WHERE id = ?",
            (batch_id,),
        )
        return self._to_metadata(row) if row is not None else None

    get_batch = get_by_id

    def require_by_id(self, batch_id: int) -> BatchMetadata:
        batch = self.get_by_id(batch_id)
        if batch is None:
            raise KeyError(f"Không tìm thấy batch {batch_id}.")
        return batch

    def get_by_sha256(self, sha256: str) -> BatchMetadata | None:
        row = self.database.query_one(
            "SELECT * FROM batches WHERE sha256 = ?",
            (sha256,),
        )
        return self._to_metadata(row) if row is not None else None

    find_by_sha256 = get_by_sha256

    def update_batch(self, batch_id: int, **changes: Any) -> BatchMetadata:
        if not changes:
            return self.require_by_id(batch_id)
        unknown = set(changes).difference(_UPDATABLE_COLUMNS)
        if unknown:
            raise ValueError(
                f"Không được cập nhật các cột: {', '.join(sorted(unknown))}."
            )
        normalized: dict[str, Any] = {}
        for key, value in changes.items():
            if key == "status":
                value = BatchStatus(value).value
            elif key.endswith("_path"):
                value = _path_text(value)
            normalized[key] = value
        assignments = ", ".join(f"{column} = ?" for column in normalized)
        parameters = [*normalized.values(), batch_id]
        with self.database.transaction(immediate=True) as connection:
            cursor = connection.execute(
                f"UPDATE batches SET {assignments} WHERE id = ?",
                parameters,
            )
            if cursor.rowcount != 1:
                raise KeyError(f"Không tìm thấy batch {batch_id}.")
            row = connection.execute(
                "SELECT * FROM batches WHERE id = ?", (batch_id,)
            ).fetchone()
        if row is None:
            raise RuntimeError("Không thể đọc lại batch vừa cập nhật.")
        return self._to_metadata(row)

    def update_validation(
        self,
        batch_id: int,
        summary: ValidationSummary,
        *,
        error_count: int | None = None,
    ) -> BatchMetadata:
        return self.update_batch(
            batch_id,
            row_count=summary.total_rows,
            valid_count=summary.valid_count,
            warning_count=summary.warning_count,
            error_count=summary.error_count if error_count is None else error_count,
            total_amount=summary.total_amount,
        )

    def mark_opened(
        self,
        batch_id: int,
        *,
        status: BatchStatus | None = None,
    ) -> BatchMetadata:
        changes: dict[str, Any] = {"last_opened_at": local_now_iso()}
        if status is not None:
            changes["status"] = status
        return self.update_batch(batch_id, **changes)

    def mark_saved(
        self,
        batch_id: int,
        summary: ValidationSummary,
        *,
        status: BatchStatus = BatchStatus.REVIEWING,
        error_count: int | None = None,
    ) -> BatchMetadata:
        return self.update_batch(
            batch_id,
            status=status,
            last_saved_at=local_now_iso(),
            row_count=summary.total_rows,
            valid_count=summary.valid_count,
            warning_count=summary.warning_count,
            error_count=summary.error_count if error_count is None else error_count,
            total_amount=summary.total_amount,
            last_error=None,
        )

    def mark_ready(
        self,
        batch_id: int,
        ready_path: str | Path,
        summary: ValidationSummary,
    ) -> BatchMetadata:
        timestamp = local_now_iso()
        return self.update_batch(
            batch_id,
            status=BatchStatus.READY,
            ready_path=ready_path,
            confirmed_at=timestamp,
            last_saved_at=timestamp,
            row_count=summary.total_rows,
            valid_count=summary.valid_count,
            warning_count=summary.warning_count,
            error_count=summary.error_count,
            total_amount=summary.total_amount,
            last_error=None,
        )

    def list_batches(
        self,
        *,
        search: str | None = None,
        status: BatchStatus | str | None = None,
        limit: int | None = None,
    ) -> list[BatchMetadata]:
        where: list[str] = []
        parameters: list[Any] = []
        if search:
            where.append("(source_filename LIKE ? OR status LIKE ?)")
            escaped = (
                search.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
            )
            term = f"%{escaped}%"
            where[-1] = "(source_filename LIKE ? ESCAPE '\\' OR status LIKE ? ESCAPE '\\')"
            parameters.extend((term, term.upper()))
        if status is not None:
            where.append("status = ?")
            parameters.append(BatchStatus(status).value)
        sql = "SELECT * FROM batches"
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY received_at DESC, id DESC"
        if limit is not None:
            if limit < 0:
                raise ValueError("Giới hạn danh sách không được âm.")
            sql += " LIMIT ?"
            parameters.append(limit)
        return [
            self._to_metadata(row)
            for row in self.database.query_all(sql, parameters)
        ]

    def list_ready_batches(self) -> list[BatchMetadata]:
        rows = self.database.query_all(
            """
            SELECT * FROM batches
            WHERE status = 'READY' AND ready_path IS NOT NULL
            ORDER BY confirmed_at DESC, id DESC
            """
        )
        return [self._to_metadata(row) for row in rows]

    list_ready = list_ready_batches

    def get_latest_ready(self) -> BatchMetadata | None:
        row = self.database.query_one(
            """
            SELECT * FROM batches
            WHERE status = 'READY' AND ready_path IS NOT NULL
            ORDER BY confirmed_at DESC, id DESC
            LIMIT 1
            """
        )
        return self._to_metadata(row) if row is not None else None

    def list_recoverable(self) -> list[BatchMetadata]:
        rows = self.database.query_all(
            """
            SELECT * FROM batches
            WHERE status IN ('REVIEWING', 'RECEIVED')
            ORDER BY received_at DESC, id DESC
            """
        )
        return [self._to_metadata(row) for row in rows]

    def set_app_state(self, key: str, value: str | None) -> None:
        if value is None:
            self.database.execute("DELETE FROM app_state WHERE key = ?", (key,))
            return
        self.database.execute(
            """
            INSERT INTO app_state(key, value) VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )

    def get_app_state(self, key: str, default: str | None = None) -> str | None:
        row = self.database.query_one(
            "SELECT value FROM app_state WHERE key = ?",
            (key,),
        )
        if row is None:
            return default
        return row["value"]

    def set_active_batch_id(self, batch_id: int | None) -> None:
        self.set_app_state(
            APP_STATE_ACTIVE_BATCH_ID,
            None if batch_id is None else str(batch_id),
        )

    def get_active_batch_id(self) -> int | None:
        value = self.get_app_state(APP_STATE_ACTIVE_BATCH_ID)
        if value is None:
            return None
        try:
            return int(value)
        except ValueError:
            self.set_active_batch_id(None)
            return None

    def restore_active_batch(self) -> BatchMetadata | None:
        active_id = self.get_active_batch_id()
        if active_id is not None:
            active = self.get_by_id(active_id)
            if (
                active is not None
                and active.status in {BatchStatus.RECEIVED, BatchStatus.REVIEWING}
                and active.working_path.is_file()
            ):
                return active
        for candidate in self.list_recoverable():
            if candidate.working_path.is_file():
                self.set_active_batch_id(candidate.id)
                return candidate
        self.set_active_batch_id(None)
        return None

    def close(self) -> None:
        self.database.close()

    @staticmethod
    def _to_metadata(row: sqlite3.Row) -> BatchMetadata:
        source_path = row["source_output_path"]
        ready_path = row["ready_path"]
        return BatchMetadata(
            id=int(row["id"]),
            source_filename=str(row["source_filename"]),
            source_output_path=Path(source_path) if source_path else None,
            original_archive_path=Path(row["original_archive_path"]),
            working_path=Path(row["working_path"]),
            ready_path=Path(ready_path) if ready_path else None,
            sha256=str(row["sha256"]),
            status=BatchStatus(row["status"]),
            received_at=str(row["received_at"]),
            last_opened_at=row["last_opened_at"],
            last_saved_at=row["last_saved_at"],
            confirmed_at=row["confirmed_at"],
            row_count=int(row["row_count"]),
            valid_count=int(row["valid_count"]),
            warning_count=int(row["warning_count"]),
            error_count=int(row["error_count"]),
            total_amount=int(row["total_amount"]),
            last_error=row["last_error"],
            source_kind=str(row["source_kind"]),
            reconciliation_group_id=(
                int(row["reconciliation_group_id"])
                if row["reconciliation_group_id"] is not None
                else None
            ),
        )
