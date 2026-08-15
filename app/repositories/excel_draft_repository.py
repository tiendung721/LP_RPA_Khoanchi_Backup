"""Persistence for the latest Excel conflict choices per source file."""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping

from app.database import Database


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _json_text(value: Mapping[str, Any]) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    )


@dataclass(frozen=True, slots=True)
class ExcelDraftRecord:
    source_file_key: str
    operation: str
    context: dict[str, Any]
    payload: dict[str, Any]
    status: str
    run_id: int | None
    created_at: str
    updated_at: str
    completed_at: str | None
    last_error: str | None

    @property
    def context_key(self) -> str:
        """Compatibility alias for callers created before schema version 15."""

        return self.source_file_key


class ExcelDraftRepository:
    def __init__(self, database: Database | str | Path) -> None:
        self.database = (
            database if isinstance(database, Database) else Database(database)
        )

    def get_latest(self, source_file_key: str) -> ExcelDraftRecord | None:
        row = self.database.query_one(
            """
            SELECT * FROM excel_resolution_latest
            WHERE source_file_key = ?
            """,
            (source_file_key,),
        )
        return self._to_record(row) if row is not None else None

    get_unfinished = get_latest

    def save(
        self,
        *,
        source_file_key: str,
        operation: str,
        context: Mapping[str, Any],
        payload: Mapping[str, Any],
        run_id: int | None = None,
    ) -> ExcelDraftRecord:
        now = _now_iso()
        with self.database.transaction(immediate=True) as connection:
            connection.execute(
                """
                INSERT INTO excel_resolution_latest (
                    source_file_key, operation, context_json, payload_json,
                    status, run_id, created_at, updated_at, completed_at, last_error
                ) VALUES (?, ?, ?, ?, 'ACTIVE', ?, ?, ?, NULL, NULL)
                ON CONFLICT(source_file_key) DO UPDATE SET
                    operation = excluded.operation,
                    context_json = excluded.context_json,
                    payload_json = excluded.payload_json,
                    status = 'ACTIVE',
                    run_id = excluded.run_id,
                    created_at = CASE
                        WHEN excel_resolution_latest.run_id IS excluded.run_id
                        THEN excel_resolution_latest.created_at
                        ELSE excluded.created_at
                    END,
                    updated_at = excluded.updated_at,
                    completed_at = NULL,
                    last_error = NULL
                """,
                (
                    source_file_key,
                    operation,
                    _json_text(dict(context)),
                    _json_text(dict(payload)),
                    run_id,
                    now,
                    now,
                ),
            )
            row = connection.execute(
                "SELECT * FROM excel_resolution_latest WHERE source_file_key = ?",
                (source_file_key,),
            ).fetchone()
        if row is None:
            raise RuntimeError("Không thể đọc lại bản xử lý Excel vừa lưu.")
        return self._to_record(row)

    def mark_failed(self, source_file_key: str, error: object) -> None:
        self.database.execute(
            """
            UPDATE excel_resolution_latest
            SET status = 'FAILED', updated_at = ?, last_error = ?
            WHERE source_file_key = ?
            """,
            (_now_iso(), str(error), source_file_key),
        )

    def mark_succeeded(self, source_file_key: str) -> None:
        now = _now_iso()
        self.database.execute(
            """
            UPDATE excel_resolution_latest
            SET status = 'SUCCEEDED', updated_at = ?, completed_at = ?,
                last_error = NULL
            WHERE source_file_key = ?
            """,
            (now, now, source_file_key),
        )

    mark_completed = mark_succeeded

    def mark_cancelled(self, source_file_key: str) -> None:
        now = _now_iso()
        self.database.execute(
            """
            UPDATE excel_resolution_latest
            SET status = 'CANCELLED', updated_at = ?, completed_at = ?,
                last_error = NULL
            WHERE source_file_key = ?
            """,
            (now, now, source_file_key),
        )

    @staticmethod
    def _to_record(row: Any) -> ExcelDraftRecord:
        return ExcelDraftRecord(
            source_file_key=str(row["source_file_key"]),
            operation=str(row["operation"]),
            context=json.loads(str(row["context_json"])),
            payload=json.loads(str(row["payload_json"])),
            status=str(row["status"]),
            run_id=None if row["run_id"] is None else int(row["run_id"]),
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
            completed_at=row["completed_at"],
            last_error=row["last_error"],
        )


__all__ = ["ExcelDraftRecord", "ExcelDraftRepository"]
