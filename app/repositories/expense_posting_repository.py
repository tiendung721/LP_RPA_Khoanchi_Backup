"""Repository cho lịch sử từng khoản chi được đối chiếu/ghi vào workbook BK."""

from __future__ import annotations

import sqlite3
from dataclasses import asdict, dataclass, is_dataclass
from enum import Enum
from pathlib import Path
from typing import Any, Iterable, Mapping

from app.database import Database
from app.repositories.excel_run_repository import (
    _enum_text,
    _json_text,
    _json_value,
    local_now_iso,
)

POSTING_ITEM_STATUSES = frozenset(
    {
        "PLANNED",
        "POSTED",
        "ALREADY_EXISTS",
        "USER_SKIPPED",
        "NOT_MATCHED",
        "UNRESOLVED",
        "FAILED",
    }
)
SUCCESSFUL_POSTING_STATUSES = frozenset({"POSTED", "ALREADY_EXISTS"})

_UPDATABLE_COLUMNS = frozenset(
    {
        "container",
        "bl",
        "fee_original",
        "fee_selected",
        "rule",
        "amount",
        "sheet_name",
        "target_row",
        "target_column",
        "target_cell",
        "value_before",
        "value_after",
        "action",
        "invoice_no",
        "invoice_selected",
        "invoice_target_column",
        "invoice_target_cell",
        "invoice_value_before",
        "invoice_value_after",
        "invoice_action",
        "carrier_source",
        "carrier_effective",
        "carrier_group",
        "carrier_target_column",
        "carrier_target_cell",
        "carrier_value_before",
        "carrier_value_after",
        "carrier_action",
        "status",
    }
)
_VALUE_COLUMNS = frozenset(
    {
        "value_before",
        "value_after",
        "invoice_value_before",
        "invoice_value_after",
        "carrier_value_before",
        "carrier_value_after",
    }
)
_POSITIVE_COLUMNS = frozenset(
    {
        "target_row",
        "target_column",
        "invoice_target_column",
        "carrier_target_column",
    }
)


def _posting_status(value: object) -> str:
    status = _enum_text(value, field_name="Trạng thái khoản chi")
    if status not in POSTING_ITEM_STATUSES:
        raise ValueError(f"Trạng thái khoản chi không hợp lệ: {status}.")
    return status


def _required_positive_int(value: object, *, field_name: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ValueError(f"{field_name} phải là số nguyên dương.")
    return value


def _required_nonnegative_int(value: object, *, field_name: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ValueError(f"{field_name} phải là số nguyên không âm.")
    return value


def _optional_text(value: object | None) -> str | None:
    if value is None:
        return None
    raw = value.value if isinstance(value, Enum) else value
    if not isinstance(raw, str):
        raise ValueError("Giá trị văn bản lịch sử khoản chi không hợp lệ.")
    return raw


@dataclass(frozen=True, slots=True)
class ExpensePostingItemRecord:
    id: int
    run_id: int
    batch_id: int
    batch_hash: str
    source_item_index: int
    container: str | None
    bl: str | None
    fee_original: str
    fee_selected: str | None
    rule: str | None
    amount: int
    sheet_name: str | None
    target_row: int | None
    target_column: int | None
    target_cell: str | None
    value_before: Any
    value_after: Any
    action: str | None
    invoice_no: str | None
    invoice_selected: str | None
    invoice_target_column: int | None
    invoice_target_cell: str | None
    invoice_value_before: Any
    invoice_value_after: Any
    invoice_action: str | None
    carrier_source: str | None
    carrier_effective: str | None
    carrier_group: str | None
    carrier_target_column: int | None
    carrier_target_cell: str | None
    carrier_value_before: Any
    carrier_value_after: Any
    carrier_action: str | None
    status: str
    created_at: str

    @property
    def item_id(self) -> int:
        return self.id


@dataclass(frozen=True, slots=True)
class ExpenseCarryForwardRecord:
    id: int
    workbook_path: str
    source_sheet: str
    source_row: int
    source_sqt: int
    container: str
    source_signature: str
    target_sheet: str
    target_row: int
    created_at: str
    updated_at: str


class ExpensePostingRepository:
    def __init__(self, database: Database | str | Path) -> None:
        self.database = (
            database if isinstance(database, Database) else Database(database)
        )

    def create_item(
        self,
        *,
        run_id: int,
        batch_id: int,
        batch_hash: str,
        source_item_index: int,
        fee_original: object,
        amount: int,
        status: object = "PLANNED",
        container: str | None = None,
        bl: str | None = None,
        fee_selected: object | None = None,
        rule: object | None = None,
        sheet_name: str | None = None,
        target_row: int | None = None,
        target_column: int | None = None,
        target_cell: str | None = None,
        value_before: object | None = None,
        value_after: object | None = None,
        action: object | None = None,
        invoice_no: str | None = None,
        invoice_selected: str | None = None,
        invoice_target_column: int | None = None,
        invoice_target_cell: str | None = None,
        invoice_value_before: object | None = None,
        invoice_value_after: object | None = None,
        invoice_action: object | None = None,
        carrier_source: str | None = None,
        carrier_effective: str | None = None,
        carrier_group: str | None = None,
        carrier_target_column: int | None = None,
        carrier_target_cell: str | None = None,
        carrier_value_before: object | None = None,
        carrier_value_after: object | None = None,
        carrier_action: object | None = None,
        created_at: str | None = None,
    ) -> ExpensePostingItemRecord:
        payload = self._normalize_create_payload(
            {
                "run_id": run_id,
                "batch_id": batch_id,
                "batch_hash": batch_hash,
                "source_item_index": source_item_index,
                "container": container,
                "bl": bl,
                "fee_original": fee_original,
                "fee_selected": fee_selected,
                "rule": rule,
                "amount": amount,
                "sheet_name": sheet_name,
                "target_row": target_row,
                "target_column": target_column,
                "target_cell": target_cell,
                "value_before": value_before,
                "value_after": value_after,
                "action": action,
                "invoice_no": invoice_no,
                "invoice_selected": invoice_selected,
                "invoice_target_column": invoice_target_column,
                "invoice_target_cell": invoice_target_cell,
                "invoice_value_before": invoice_value_before,
                "invoice_value_after": invoice_value_after,
                "invoice_action": invoice_action,
                "carrier_source": carrier_source,
                "carrier_effective": carrier_effective,
                "carrier_group": carrier_group,
                "carrier_target_column": carrier_target_column,
                "carrier_target_cell": carrier_target_cell,
                "carrier_value_before": carrier_value_before,
                "carrier_value_after": carrier_value_after,
                "carrier_action": carrier_action,
                "status": status,
                "created_at": created_at or local_now_iso(),
            }
        )
        with self.database.transaction(immediate=True) as connection:
            item_id = self._insert(connection, payload)
            row = connection.execute(
                "SELECT * FROM expense_posting_items WHERE id = ?",
                (item_id,),
            ).fetchone()
        if row is None:
            raise RuntimeError("Không thể đọc lại lịch sử khoản chi vừa tạo.")
        return self._to_record(row)

    def create_items(
        self,
        items: Iterable[Mapping[str, Any] | object],
        *,
        run_id: int | None = None,
        batch_id: int | None = None,
        batch_hash: str | None = None,
        status: object | None = None,
    ) -> list[ExpensePostingItemRecord]:
        payloads: list[dict[str, Any]] = []
        for item in items:
            if isinstance(item, Mapping):
                payload = dict(item)
            elif is_dataclass(item) and not isinstance(item, type):
                payload = asdict(item)
            else:
                raise TypeError("Mỗi khoản chi phải là Mapping hoặc dataclass.")
            defaults = {
                "run_id": run_id,
                "batch_id": batch_id,
                "batch_hash": batch_hash,
            }
            if status is not None:
                defaults["status"] = status
            for key, value in defaults.items():
                if key not in payload and value is not None:
                    payload[key] = value
            payload.setdefault("status", "PLANNED")
            payload.setdefault("created_at", local_now_iso())
            payloads.append(self._normalize_create_payload(payload))
        if not payloads:
            return []

        ids: list[int] = []
        with self.database.transaction(immediate=True) as connection:
            for payload in payloads:
                ids.append(self._insert(connection, payload))
            placeholders = ", ".join("?" for _ in ids)
            rows = connection.execute(
                f"""
                SELECT * FROM expense_posting_items
                WHERE id IN ({placeholders})
                ORDER BY id
                """,
                ids,
            ).fetchall()
        return [self._to_record(row) for row in rows]

    save_items = create_items
    record_items = create_items

    def get_carry_forward(
        self,
        *,
        workbook_path: str | Path,
        source_sheet: str,
        source_row: int,
        source_sqt: int,
        container: str,
        target_sheet: str,
    ) -> ExpenseCarryForwardRecord | None:
        row = self.database.query_one(
            """
            SELECT * FROM expense_carry_forwards
            WHERE workbook_path = ? AND source_sheet = ? AND source_row = ?
              AND source_sqt = ? AND container = ? AND target_sheet = ?
            """,
            (
                str(Path(workbook_path).resolve(strict=False)),
                source_sheet,
                source_row,
                source_sqt,
                container,
                target_sheet,
            ),
        )
        return self._to_carry_forward(row) if row is not None else None

    def save_carry_forward(
        self,
        *,
        workbook_path: str | Path,
        source_sheet: str,
        source_row: int,
        source_sqt: int,
        container: str,
        source_signature: str,
        target_sheet: str,
        target_row: int,
    ) -> ExpenseCarryForwardRecord:
        if source_row <= 0 or source_sqt <= 0 or target_row <= 0:
            raise ValueError("Dòng nguồn, SQT và dòng đích phải là số nguyên dương.")
        if not container or not source_signature:
            raise ValueError("Container và chữ ký dòng nguồn không được để trống.")
        now = local_now_iso()
        resolved_path = str(Path(workbook_path).resolve(strict=False))
        with self.database.transaction(immediate=True) as connection:
            connection.execute(
                """
                INSERT INTO expense_carry_forwards (
                    workbook_path, source_sheet, source_row, source_sqt,
                    container, source_signature, target_sheet, target_row,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT (
                    workbook_path, source_sheet, source_row, source_sqt,
                    container, target_sheet
                ) DO UPDATE SET
                    source_signature = excluded.source_signature,
                    target_row = excluded.target_row,
                    updated_at = excluded.updated_at
                """,
                (
                    resolved_path,
                    source_sheet,
                    source_row,
                    source_sqt,
                    container,
                    source_signature,
                    target_sheet,
                    target_row,
                    now,
                    now,
                ),
            )
            row = connection.execute(
                """
                SELECT * FROM expense_carry_forwards
                WHERE workbook_path = ? AND source_sheet = ? AND source_row = ?
                  AND source_sqt = ? AND container = ? AND target_sheet = ?
                """,
                (
                    resolved_path,
                    source_sheet,
                    source_row,
                    source_sqt,
                    container,
                    target_sheet,
                ),
            ).fetchone()
        if row is None:
            raise RuntimeError("Không thể đọc lại ánh xạ dòng kế hoạch vừa lưu.")
        return self._to_carry_forward(row)

    @staticmethod
    def _to_carry_forward(row: sqlite3.Row) -> ExpenseCarryForwardRecord:
        return ExpenseCarryForwardRecord(
            id=int(row["id"]),
            workbook_path=str(row["workbook_path"]),
            source_sheet=str(row["source_sheet"]),
            source_row=int(row["source_row"]),
            source_sqt=int(row["source_sqt"]),
            container=str(row["container"]),
            source_signature=str(row["source_signature"]),
            target_sheet=str(row["target_sheet"]),
            target_row=int(row["target_row"]),
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
        )

    def get_by_id(self, item_id: int) -> ExpensePostingItemRecord | None:
        row = self.database.query_one(
            "SELECT * FROM expense_posting_items WHERE id = ?",
            (item_id,),
        )
        return self._to_record(row) if row is not None else None

    def require_by_id(self, item_id: int) -> ExpensePostingItemRecord:
        record = self.get_by_id(item_id)
        if record is None:
            raise KeyError(f"Không tìm thấy lịch sử khoản chi {item_id}.")
        return record

    def update_item(
        self,
        item_id: int,
        **changes: Any,
    ) -> ExpensePostingItemRecord:
        if not changes:
            return self.require_by_id(item_id)
        unknown = set(changes).difference(_UPDATABLE_COLUMNS)
        if unknown:
            raise ValueError(
                f"Không được cập nhật các cột: {', '.join(sorted(unknown))}."
            )
        normalized: dict[str, Any] = {}
        for key, value in changes.items():
            if key == "status":
                value = _posting_status(value)
            elif key in _VALUE_COLUMNS:
                value = _json_text(value)
            elif key in _POSITIVE_COLUMNS:
                if value is not None:
                    value = _required_positive_int(value, field_name=key)
            elif key == "amount":
                value = _required_nonnegative_int(value, field_name="Số tiền")
            elif key in {
                "container",
                "bl",
                "fee_original",
                "fee_selected",
                "rule",
                "sheet_name",
                "target_cell",
                "action",
                "invoice_no",
                "invoice_selected",
                "invoice_target_cell",
                "invoice_action",
            }:
                value = _optional_text(value)
                if key == "fee_original" and not value:
                    raise ValueError("Mã phí gốc không được để trống.")
            normalized[key] = value
        assignments = ", ".join(f"{column} = ?" for column in normalized)
        with self.database.transaction(immediate=True) as connection:
            cursor = connection.execute(
                f"UPDATE expense_posting_items SET {assignments} WHERE id = ?",
                (*normalized.values(), item_id),
            )
            if cursor.rowcount != 1:
                raise KeyError(f"Không tìm thấy lịch sử khoản chi {item_id}.")
            row = connection.execute(
                "SELECT * FROM expense_posting_items WHERE id = ?",
                (item_id,),
            ).fetchone()
        if row is None:
            raise RuntimeError("Không thể đọc lại lịch sử khoản chi vừa cập nhật.")
        return self._to_record(row)

    def list_by_run(self, run_id: int) -> list[ExpensePostingItemRecord]:
        rows = self.database.query_all(
            """
            SELECT * FROM expense_posting_items
            WHERE run_id = ?
            ORDER BY source_item_index, id
            """,
            (run_id,),
        )
        return [self._to_record(row) for row in rows]

    def list_by_batch_hash(
        self,
        batch_hash: str,
        *,
        statuses: Iterable[object] | object | None = None,
    ) -> list[ExpensePostingItemRecord]:
        parameters: list[Any] = [batch_hash]
        sql = "SELECT * FROM expense_posting_items WHERE batch_hash = ?"
        if statuses is not None:
            values = (
                [statuses]
                if isinstance(statuses, (str, Enum))
                else list(statuses)  # type: ignore[arg-type]
            )
            normalized_statuses = [_posting_status(value) for value in values]
            if not normalized_statuses:
                return []
            placeholders = ", ".join("?" for _ in normalized_statuses)
            sql += f" AND status IN ({placeholders})"
            parameters.extend(normalized_statuses)
        sql += " ORDER BY source_item_index, id"
        return [
            self._to_record(row)
            for row in self.database.query_all(sql, parameters)
        ]

    def list_successful_by_batch_hash(
        self,
        batch_hash: str,
    ) -> list[ExpensePostingItemRecord]:
        return self.list_by_batch_hash(
            batch_hash,
            statuses=SUCCESSFUL_POSTING_STATUSES,
        )

    def successful_source_indices(self, batch_hash: str) -> set[int]:
        rows = self.database.query_all(
            """
            SELECT source_item_index
            FROM expense_posting_items
            WHERE batch_hash = ?
              AND status IN ('POSTED', 'ALREADY_EXISTS')
            """,
            (batch_hash,),
        )
        return {int(row["source_item_index"]) for row in rows}

    def latest_successful_items(
        self,
        batch_hash: str,
    ) -> list[ExpensePostingItemRecord]:
        """Trả về lần nhập thành công mới nhất của từng dòng nguồn."""

        rows = self.database.query_all(
            """
            SELECT item.*
            FROM expense_posting_items AS item
            WHERE item.batch_hash = ?
              AND item.status IN ('POSTED', 'ALREADY_EXISTS')
              AND NOT EXISTS (
                    SELECT 1
                    FROM expense_posting_items AS newer
                    WHERE newer.batch_hash = item.batch_hash
                      AND newer.source_item_index = item.source_item_index
                      AND newer.status IN ('POSTED', 'ALREADY_EXISTS')
                      AND newer.id > item.id
              )
            ORDER BY item.source_item_index
            """,
            (batch_hash,),
        )
        return [self._to_record(row) for row in rows]

    get_latest_successful_items = latest_successful_items

    get_successful_source_indices = successful_source_indices
    posted_source_indices = successful_source_indices

    def is_source_item_posted(
        self,
        batch_hash: str,
        source_item_index: int,
    ) -> bool:
        row = self.database.query_one(
            """
            SELECT 1
            FROM expense_posting_items
            WHERE batch_hash = ?
              AND source_item_index = ?
              AND status IN ('POSTED', 'ALREADY_EXISTS')
            LIMIT 1
            """,
            (batch_hash, source_item_index),
        )
        return row is not None

    def batch_has_successful_items(self, batch_hash: str) -> bool:
        row = self.database.query_one(
            """
            SELECT 1
            FROM expense_posting_items
            WHERE batch_hash = ?
              AND status IN ('POSTED', 'ALREADY_EXISTS')
            LIMIT 1
            """,
            (batch_hash,),
        )
        return row is not None

    has_successful_items = batch_has_successful_items

    def get_latest_successful_for_batch(
        self,
        batch_hash: str,
    ) -> ExpensePostingItemRecord | None:
        row = self.database.query_one(
            """
            SELECT item.*
            FROM expense_posting_items AS item
            JOIN excel_runs AS run ON run.id = item.run_id
            WHERE item.batch_hash = ?
              AND item.status IN ('POSTED', 'ALREADY_EXISTS')
            ORDER BY COALESCE(run.completed_at, run.started_at) DESC, item.id DESC
            LIMIT 1
            """,
            (batch_hash,),
        )
        return self._to_record(row) if row is not None else None

    @staticmethod
    def _insert(
        connection: sqlite3.Connection,
        payload: Mapping[str, Any],
    ) -> int:
        cursor = connection.execute(
            """
            INSERT INTO expense_posting_items (
                run_id, batch_id, batch_hash, source_item_index, container, bl,
                fee_original, fee_selected, rule, amount, sheet_name,
                target_row, target_column, target_cell, value_before,
                value_after, action, invoice_no, invoice_selected,
                invoice_target_column, invoice_target_cell,
                invoice_value_before, invoice_value_after, invoice_action,
                carrier_source, carrier_effective, carrier_group,
                carrier_target_column, carrier_target_cell,
                carrier_value_before, carrier_value_after, carrier_action,
                status, created_at
            ) VALUES (
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
            )
            """,
            tuple(
                payload[column]
                for column in (
                    "run_id",
                    "batch_id",
                    "batch_hash",
                    "source_item_index",
                    "container",
                    "bl",
                    "fee_original",
                    "fee_selected",
                    "rule",
                    "amount",
                    "sheet_name",
                    "target_row",
                    "target_column",
                    "target_cell",
                    "value_before",
                    "value_after",
                    "action",
                    "invoice_no",
                    "invoice_selected",
                    "invoice_target_column",
                    "invoice_target_cell",
                    "invoice_value_before",
                    "invoice_value_after",
                    "invoice_action",
                    "carrier_source",
                    "carrier_effective",
                    "carrier_group",
                    "carrier_target_column",
                    "carrier_target_cell",
                    "carrier_value_before",
                    "carrier_value_after",
                    "carrier_action",
                    "status",
                    "created_at",
                )
            ),
        )
        return int(cursor.lastrowid)

    @staticmethod
    def _normalize_create_payload(payload: Mapping[str, Any]) -> dict[str, Any]:
        required = {
            "run_id",
            "batch_id",
            "batch_hash",
            "source_item_index",
            "fee_original",
            "amount",
            "status",
            "created_at",
        }
        missing = required.difference(payload)
        if missing:
            raise ValueError(
                f"Thiếu trường lịch sử khoản chi: {', '.join(sorted(missing))}."
            )
        run_id = _required_positive_int(payload["run_id"], field_name="run_id")
        batch_id = _required_positive_int(
            payload["batch_id"],
            field_name="batch_id",
        )
        source_item_index = _required_nonnegative_int(
            payload["source_item_index"],
            field_name="source_item_index",
        )
        amount = _required_nonnegative_int(payload["amount"], field_name="Số tiền")
        batch_hash = _optional_text(payload["batch_hash"])
        fee_original = _optional_text(payload["fee_original"])
        created_at = _optional_text(payload["created_at"])
        if not batch_hash or not batch_hash.strip():
            raise ValueError("Hash batch không được để trống.")
        if not fee_original or not fee_original.strip():
            raise ValueError("Mã phí gốc không được để trống.")
        if not created_at:
            raise ValueError("Thời gian tạo lịch sử không được để trống.")

        target_row = payload.get("target_row")
        target_column = payload.get("target_column")
        invoice_target_column = payload.get("invoice_target_column")
        carrier_target_column = payload.get("carrier_target_column")
        if target_row is not None:
            target_row = _required_positive_int(target_row, field_name="target_row")
        if target_column is not None:
            target_column = _required_positive_int(
                target_column,
                field_name="target_column",
            )
        if invoice_target_column is not None:
            invoice_target_column = _required_positive_int(
                invoice_target_column,
                field_name="invoice_target_column",
            )
        if carrier_target_column is not None:
            carrier_target_column = _required_positive_int(
                carrier_target_column,
                field_name="carrier_target_column",
            )
        return {
            "run_id": run_id,
            "batch_id": batch_id,
            "batch_hash": batch_hash,
            "source_item_index": source_item_index,
            "container": _optional_text(payload.get("container")),
            "bl": _optional_text(payload.get("bl")),
            "fee_original": fee_original,
            "fee_selected": _optional_text(payload.get("fee_selected")),
            "rule": _optional_text(payload.get("rule")),
            "amount": amount,
            "sheet_name": _optional_text(payload.get("sheet_name")),
            "target_row": target_row,
            "target_column": target_column,
            "target_cell": _optional_text(payload.get("target_cell")),
            "value_before": _json_text(payload.get("value_before")),
            "value_after": _json_text(payload.get("value_after")),
            "action": _optional_text(payload.get("action")),
            "invoice_no": _optional_text(payload.get("invoice_no")),
            "invoice_selected": _optional_text(payload.get("invoice_selected")),
            "invoice_target_column": invoice_target_column,
            "invoice_target_cell": _optional_text(payload.get("invoice_target_cell")),
            "invoice_value_before": _json_text(payload.get("invoice_value_before")),
            "invoice_value_after": _json_text(payload.get("invoice_value_after")),
            "invoice_action": _optional_text(payload.get("invoice_action")),
            "carrier_source": _optional_text(payload.get("carrier_source")),
            "carrier_effective": _optional_text(payload.get("carrier_effective")),
            "carrier_group": _optional_text(payload.get("carrier_group")),
            "carrier_target_column": carrier_target_column,
            "carrier_target_cell": _optional_text(payload.get("carrier_target_cell")),
            "carrier_value_before": _json_text(payload.get("carrier_value_before")),
            "carrier_value_after": _json_text(payload.get("carrier_value_after")),
            "carrier_action": _optional_text(payload.get("carrier_action")),
            "status": _posting_status(payload["status"]),
            "created_at": created_at,
        }

    @staticmethod
    def _to_record(row: sqlite3.Row) -> ExpensePostingItemRecord:
        return ExpensePostingItemRecord(
            id=int(row["id"]),
            run_id=int(row["run_id"]),
            batch_id=int(row["batch_id"]),
            batch_hash=str(row["batch_hash"]),
            source_item_index=int(row["source_item_index"]),
            container=row["container"],
            bl=row["bl"],
            fee_original=str(row["fee_original"]),
            fee_selected=row["fee_selected"],
            rule=row["rule"],
            amount=int(row["amount"]),
            sheet_name=row["sheet_name"],
            target_row=(
                int(row["target_row"]) if row["target_row"] is not None else None
            ),
            target_column=(
                int(row["target_column"])
                if row["target_column"] is not None
                else None
            ),
            target_cell=row["target_cell"],
            value_before=_json_value(row["value_before"]),
            value_after=_json_value(row["value_after"]),
            action=row["action"],
            invoice_no=row["invoice_no"],
            invoice_selected=row["invoice_selected"],
            invoice_target_column=(
                int(row["invoice_target_column"])
                if row["invoice_target_column"] is not None
                else None
            ),
            invoice_target_cell=row["invoice_target_cell"],
            invoice_value_before=_json_value(row["invoice_value_before"]),
            invoice_value_after=_json_value(row["invoice_value_after"]),
            invoice_action=row["invoice_action"],
            carrier_source=row["carrier_source"],
            carrier_effective=row["carrier_effective"],
            carrier_group=row["carrier_group"],
            carrier_target_column=(
                int(row["carrier_target_column"])
                if row["carrier_target_column"] is not None
                else None
            ),
            carrier_target_cell=row["carrier_target_cell"],
            carrier_value_before=_json_value(row["carrier_value_before"]),
            carrier_value_after=_json_value(row["carrier_value_after"]),
            carrier_action=row["carrier_action"],
            status=str(row["status"]),
            created_at=str(row["created_at"]),
        )


__all__ = [
    "POSTING_ITEM_STATUSES",
    "SUCCESSFUL_POSTING_STATUSES",
    "ExpenseCarryForwardRecord",
    "ExpensePostingItemRecord",
    "ExpensePostingRepository",
]
