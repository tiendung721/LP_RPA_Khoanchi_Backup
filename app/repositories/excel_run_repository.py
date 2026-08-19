"""Repository cho lịch sử các lần phân tích và ghi workbook Excel."""

from __future__ import annotations

import json
import sqlite3
from dataclasses import asdict, dataclass, is_dataclass
from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from pathlib import Path
from typing import Any, Iterable, Mapping

from app.database import Database

EXCEL_RUN_STATUSES = frozenset(
    {
        "ANALYZING",
        "WAITING_USER",
        "APPLYING",
        "SUCCEEDED",
        "NO_CHANGES",
        "CANCELLED",
        "FAILED",
    }
)

_UPDATABLE_COLUMNS = frozenset(
    {
        "completed_at",
        "source_path",
        "target_path",
        "source_fingerprint",
        "target_fingerprint_before",
        "target_fingerprint_after",
        "sheet_name",
        "backup_path",
        "status",
        "total_items",
        "changed_items",
        "skipped_items",
        "conflict_count",
        "error_message",
        "item_outcomes",
    }
)
_PATH_COLUMNS = frozenset({"source_path", "target_path", "backup_path"})
_FINGERPRINT_COLUMNS = frozenset(
    {
        "source_fingerprint",
        "target_fingerprint_before",
        "target_fingerprint_after",
    }
)
_JSON_COLUMNS = _FINGERPRINT_COLUMNS | frozenset({"item_outcomes"})
_COUNT_COLUMNS = frozenset(
    {"total_items", "changed_items", "skipped_items", "conflict_count"}
)


def local_now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _enum_text(value: object, *, field_name: str) -> str:
    raw = value.value if isinstance(value, Enum) else value
    if not isinstance(raw, str) or not raw.strip():
        raise ValueError(f"{field_name} phải là chuỗi không rỗng.")
    return raw.strip()


def _run_status(value: object) -> str:
    status = _enum_text(value, field_name="Trạng thái lần chạy")
    if status not in EXCEL_RUN_STATUSES:
        raise ValueError(f"Trạng thái lần chạy không hợp lệ: {status}.")
    return status


def _json_default(value: object) -> object:
    if is_dataclass(value) and not isinstance(value, type):
        return asdict(value)
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return str(value)
    to_dict = getattr(value, "to_dict", None)
    if callable(to_dict):
        return to_dict()
    raise TypeError(f"Không thể lưu giá trị JSON kiểu {type(value).__name__}.")


def _json_text(value: object | None) -> str | None:
    if value is None:
        return None
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
        default=_json_default,
    )


def _json_value(value: object | None) -> Any:
    if value is None:
        return None
    try:
        return json.loads(str(value))
    except (TypeError, ValueError, json.JSONDecodeError):
        # Giữ khả năng đọc database thử nghiệm/tiền phát hành từng lưu chuỗi thô.
        return value


@dataclass(frozen=True, slots=True)
class ExcelRunRecord:
    id: int
    operation: str
    started_at: str
    completed_at: str | None
    source_path: Path | None
    target_path: Path | None
    source_fingerprint: Any
    target_fingerprint_before: Any
    target_fingerprint_after: Any
    sheet_name: str | None
    backup_path: Path | None
    status: str
    total_items: int
    changed_items: int
    skipped_items: int
    conflict_count: int
    error_message: str | None
    item_outcomes: Any

    @property
    def run_id(self) -> int:
        return self.id


class ExcelRunRepository:
    def __init__(self, database: Database | str | Path) -> None:
        self.database = (
            database if isinstance(database, Database) else Database(database)
        )

    def create_run(
        self,
        *,
        operation: object,
        status: object = "ANALYZING",
        started_at: str | None = None,
        source_path: str | Path | None = None,
        target_path: str | Path | None = None,
        source_fingerprint: object | None = None,
        target_fingerprint_before: object | None = None,
        sheet_name: str | None = None,
    ) -> ExcelRunRecord:
        operation_text = _enum_text(operation, field_name="Nghiệp vụ Excel")
        status_text = _run_status(status)
        with self.database.transaction(immediate=True) as connection:
            cursor = connection.execute(
                """
                INSERT INTO excel_runs (
                    operation, started_at, source_path, target_path,
                    source_fingerprint, target_fingerprint_before,
                    sheet_name, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    operation_text,
                    started_at or local_now_iso(),
                    None if source_path is None else str(source_path),
                    None if target_path is None else str(target_path),
                    _json_text(source_fingerprint),
                    _json_text(target_fingerprint_before),
                    sheet_name,
                    status_text,
                ),
            )
            row = connection.execute(
                "SELECT * FROM excel_runs WHERE id = ?",
                (int(cursor.lastrowid),),
            ).fetchone()
        if row is None:
            raise RuntimeError("Không thể đọc lại lịch sử Excel vừa tạo.")
        return self._to_record(row)

    def get_by_id(self, run_id: int) -> ExcelRunRecord | None:
        row = self.database.query_one(
            "SELECT * FROM excel_runs WHERE id = ?",
            (run_id,),
        )
        return self._to_record(row) if row is not None else None

    get_run = get_by_id

    def require_by_id(self, run_id: int) -> ExcelRunRecord:
        record = self.get_by_id(run_id)
        if record is None:
            raise KeyError(f"Không tìm thấy lần chạy Excel {run_id}.")
        return record

    def update_run(self, run_id: int, **changes: Any) -> ExcelRunRecord:
        if not changes:
            return self.require_by_id(run_id)
        unknown = set(changes).difference(_UPDATABLE_COLUMNS)
        if unknown:
            raise ValueError(
                f"Không được cập nhật các cột: {', '.join(sorted(unknown))}."
            )
        normalized: dict[str, Any] = {}
        for key, value in changes.items():
            if key == "status":
                value = _run_status(value)
            elif key in _PATH_COLUMNS:
                value = None if value is None else str(value)
            elif key in _JSON_COLUMNS:
                value = _json_text(value)
            elif key in _COUNT_COLUMNS:
                if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                    raise ValueError(f"{key} phải là số nguyên không âm.")
            normalized[key] = value
        assignments = ", ".join(f"{column} = ?" for column in normalized)
        with self.database.transaction(immediate=True) as connection:
            cursor = connection.execute(
                f"UPDATE excel_runs SET {assignments} WHERE id = ?",
                (*normalized.values(), run_id),
            )
            if cursor.rowcount != 1:
                raise KeyError(f"Không tìm thấy lần chạy Excel {run_id}.")
            row = connection.execute(
                "SELECT * FROM excel_runs WHERE id = ?",
                (run_id,),
            ).fetchone()
        if row is None:
            raise RuntimeError("Không thể đọc lại lịch sử Excel vừa cập nhật.")
        return self._to_record(row)

    def finish_run(
        self,
        run_id: int,
        *,
        status: object = "SUCCEEDED",
        completed_at: str | None = None,
        **changes: Any,
    ) -> ExcelRunRecord:
        changes["status"] = status
        changes["completed_at"] = completed_at or local_now_iso()
        return self.update_run(run_id, **changes)

    complete_run = finish_run

    def get_latest(
        self,
        *,
        operation: object | None = None,
        statuses: Iterable[object] | object | None = None,
    ) -> ExcelRunRecord | None:
        where: list[str] = []
        parameters: list[Any] = []
        if operation is not None:
            where.append("operation = ?")
            parameters.append(_enum_text(operation, field_name="Nghiệp vụ Excel"))
        if statuses is not None:
            values = (
                [statuses]
                if isinstance(statuses, (str, Enum))
                else list(statuses)  # type: ignore[arg-type]
            )
            normalized_statuses = [_run_status(value) for value in values]
            if not normalized_statuses:
                return None
            placeholders = ", ".join("?" for _ in normalized_statuses)
            where.append(f"status IN ({placeholders})")
            parameters.extend(normalized_statuses)
        sql = "SELECT * FROM excel_runs"
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY started_at DESC, id DESC LIMIT 1"
        row = self.database.query_one(sql, parameters)
        return self._to_record(row) if row is not None else None

    def list_runs(
        self,
        *,
        operation: object | None = None,
        limit: int | None = None,
    ) -> list[ExcelRunRecord]:
        parameters: list[Any] = []
        sql = "SELECT * FROM excel_runs"
        if operation is not None:
            sql += " WHERE operation = ?"
            parameters.append(_enum_text(operation, field_name="Nghiệp vụ Excel"))
        sql += " ORDER BY started_at DESC, id DESC"
        if limit is not None:
            if isinstance(limit, bool) or not isinstance(limit, int) or limit < 0:
                raise ValueError("Giới hạn danh sách phải là số nguyên không âm.")
            sql += " LIMIT ?"
            parameters.append(limit)
        return [
            self._to_record(row)
            for row in self.database.query_all(sql, parameters)
        ]

    def get_latest_successful_sync(
        self,
        *,
        operation: object | None = None,
    ) -> ExcelRunRecord | None:
        parameters: list[Any]
        if operation is None:
            operation_filter = (
                "UPPER(operation) IN ('DAILY_SYNC', 'SYNC', 'SYNC_DAILY', "
                "'DAILY WORKBOOK SYNC')"
            )
            parameters = []
        else:
            operation_filter = "operation = ?"
            parameters = [_enum_text(operation, field_name="Nghiệp vụ Excel")]
        row = self.database.query_one(
            f"""
            SELECT * FROM excel_runs
            WHERE {operation_filter}
              AND status = 'SUCCEEDED'
              AND sheet_name IS NOT NULL
            ORDER BY COALESCE(completed_at, started_at) DESC, id DESC
            LIMIT 1
            """,
            parameters,
        )
        return self._to_record(row) if row is not None else None

    def get_latest_sync_sheet(self, *, operation: object | None = None) -> str | None:
        record = self.get_latest_successful_sync(operation=operation)
        return record.sheet_name if record is not None else None

    latest_successful_sheet = get_latest_sync_sheet

    @staticmethod
    def _to_record(row: sqlite3.Row) -> ExcelRunRecord:
        return ExcelRunRecord(
            id=int(row["id"]),
            operation=str(row["operation"]),
            started_at=str(row["started_at"]),
            completed_at=row["completed_at"],
            source_path=Path(row["source_path"]) if row["source_path"] else None,
            target_path=Path(row["target_path"]) if row["target_path"] else None,
            source_fingerprint=_json_value(row["source_fingerprint"]),
            target_fingerprint_before=_json_value(
                row["target_fingerprint_before"]
            ),
            target_fingerprint_after=_json_value(row["target_fingerprint_after"]),
            sheet_name=row["sheet_name"],
            backup_path=Path(row["backup_path"]) if row["backup_path"] else None,
            status=str(row["status"]),
            total_items=int(row["total_items"]),
            changed_items=int(row["changed_items"]),
            skipped_items=int(row["skipped_items"]),
            conflict_count=int(row["conflict_count"]),
            error_message=row["error_message"],
            item_outcomes=_json_value(row["item_outcomes"]),
        )


__all__ = [
    "EXCEL_RUN_STATUSES",
    "ExcelRunRecord",
    "ExcelRunRepository",
]
