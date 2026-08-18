from __future__ import annotations

import json
from datetime import datetime
from typing import Any

from app.database import Database

from .contracts import (
    BkContainerSnapshot,
    GroupStatus,
    InvoiceContribution,
    ReconciliationGroup,
)


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


class DuplicateContributionError(ValueError):
    pass


class SeaFreightRepository:
    def __init__(self, database: Database) -> None:
        self.database = database

    def get_group(self, group_id: int) -> ReconciliationGroup | None:
        row = self.database.query_one(
            "SELECT * FROM sea_freight_reconciliation_groups WHERE id = ?",
            (group_id,),
        )
        return self._group(row) if row is not None else None

    def find_open_group(
        self, bk_path: str, bk_sheet: str, vessel_key: str, voyage_key: str
    ) -> ReconciliationGroup | None:
        row = self.database.query_one(
            """
            SELECT * FROM sea_freight_reconciliation_groups
            WHERE bk_path = ? AND bk_sheet = ? AND vessel_key = ? AND voyage_key = ?
              AND is_current = 1
              AND status NOT IN ('POSTED','CANCELLED')
            ORDER BY id DESC LIMIT 1
            """,
            (bk_path, bk_sheet, vessel_key, voyage_key),
        )
        return self._group(row) if row is not None else None

    def find_latest_group(
        self, bk_path: str, bk_sheet: str, vessel_key: str, voyage_key: str
    ) -> ReconciliationGroup | None:
        row = self.database.query_one(
            """
            SELECT * FROM sea_freight_reconciliation_groups
            WHERE bk_path = ? AND bk_sheet = ? AND vessel_key = ? AND voyage_key = ?
              AND is_current = 1
            ORDER BY id DESC LIMIT 1
            """,
            (bk_path, bk_sheet, vessel_key, voyage_key),
        )
        return self._group(row) if row is not None else None

    def list_groups(self, *, include_closed: bool = False) -> list[ReconciliationGroup]:
        where = "" if include_closed else "WHERE is_current = 1 AND status != 'CANCELLED'"
        return [
            self._group(row)
            for row in self.database.query_all(
                f"SELECT * FROM sea_freight_reconciliation_groups {where} "
                "ORDER BY updated_at DESC, id DESC"
            )
        ]

    def list_groups_for_period(self, month: int, year: int) -> list[ReconciliationGroup]:
        return [
            self._group(row)
            for row in self.database.query_all(
                """
                SELECT * FROM sea_freight_reconciliation_groups
                WHERE reconciliation_month = ? AND reconciliation_year = ?
                  AND status != 'CANCELLED'
                ORDER BY is_current DESC, vessel_voyage_raw, revision_no DESC, id DESC
                """,
                (month, year),
            )
        ]

    def upsert_snapshot(
        self,
        snapshot: BkContainerSnapshot,
        *,
        month: int | None = None,
        year: int | None = None,
        bk_issue: str | None = None,
    ) -> ReconciliationGroup:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            row = connection.execute(
                """
                SELECT * FROM sea_freight_reconciliation_groups
                WHERE bk_path = ? AND bk_sheet = ? AND vessel_key = ? AND voyage_key = ?
                  AND is_current = 1
                  AND status NOT IN ('POSTED','CANCELLED')
                ORDER BY id DESC LIMIT 1
                """,
                (
                    snapshot.bk_path,
                    snapshot.bk_sheet,
                    snapshot.vessel_key,
                    snapshot.voyage_key,
                ),
            ).fetchone()
            created = row is None
            if row is None:
                cursor = connection.execute(
                    """
                    INSERT INTO sea_freight_reconciliation_groups(
                        bk_path, bk_sheet, vessel_voyage_raw, vessel_key,
                        voyage_key, combined_key, bk_container_count,
                        received_container_count, total_amount, status,
                        bk_fingerprint, container_snapshot_hash,
                        reconciliation_month, reconciliation_year, bk_issue,
                        bk_invalid_container_count, bk_duplicate_container_count,
                        created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, 0, 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        snapshot.bk_path,
                        snapshot.bk_sheet,
                        snapshot.vessel_voyage_raw,
                        snapshot.vessel_key,
                        snapshot.voyage_key,
                        snapshot.combined_key,
                        snapshot.container_count,
                        GroupStatus.BK_NOT_FOUND.value
                        if snapshot.container_count == 0
                        else GroupStatus.PENDING.value,
                        snapshot.workbook_fingerprint,
                        snapshot.snapshot_hash,
                        month,
                        year,
                        bk_issue,
                        snapshot.invalid_container_count,
                        snapshot.duplicate_container_count,
                        timestamp,
                        timestamp,
                    ),
                )
                group_id = int(cursor.lastrowid)
            else:
                group_id = int(row["id"])
                connection.execute(
                    """
                    UPDATE sea_freight_reconciliation_groups
                    SET vessel_voyage_raw = ?, combined_key = ?,
                        bk_container_count = ?, bk_fingerprint = ?,
                        container_snapshot_hash = ?,
                        reconciliation_month = COALESCE(?, reconciliation_month),
                        reconciliation_year = COALESCE(?, reconciliation_year),
                        bk_issue = ?, bk_invalid_container_count = ?,
                        bk_duplicate_container_count = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (
                        snapshot.vessel_voyage_raw,
                        snapshot.combined_key,
                        snapshot.container_count,
                        snapshot.workbook_fingerprint,
                        snapshot.snapshot_hash,
                        month,
                        year,
                        bk_issue,
                        snapshot.invalid_container_count,
                        snapshot.duplicate_container_count,
                        timestamp,
                        group_id,
                    ),
                )
            cursor = connection.execute(
                "DELETE FROM sea_freight_group_containers WHERE group_id = ?",
                (group_id,),
            )
            for order, record in enumerate(snapshot.containers):
                connection.execute(
                    """
                    INSERT INTO sea_freight_group_containers(
                        group_id, container, source_sheet, source_row,
                        source_sqt, departure_date, allocation_order
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        group_id,
                        record.container,
                        record.source_sheet,
                        record.source_row,
                        record.source_sqt,
                        record.departure_date,
                        order,
                    ),
                )
            self._recalculate(connection, group_id, timestamp)
            self._event(connection, group_id, "GROUP_CREATED" if created else "BK_SNAPSHOT", {
                "container_count": snapshot.container_count,
                "snapshot_hash": snapshot.snapshot_hash,
            }, timestamp)
        group = self.get_group(group_id)
        assert group is not None
        return group

    def replace_group_snapshot(
        self,
        group_id: int,
        snapshot: BkContainerSnapshot,
        *,
        vessel_name: str,
        voyage_no: str,
    ) -> ReconciliationGroup:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            cursor = connection.execute(
                """
                UPDATE sea_freight_reconciliation_groups
                SET bk_path = ?, bk_sheet = ?, vessel_voyage_raw = ?,
                    vessel_key = ?, voyage_key = ?, combined_key = ?,
                    bk_container_count = ?, bk_fingerprint = ?,
                    container_snapshot_hash = ?, bk_invalid_container_count = ?,
                    bk_duplicate_container_count = ?, status = 'PENDING', updated_at = ?
                WHERE id = ? AND status NOT IN ('ALLOCATED','POSTED','CANCELLED')
                """,
                (
                    snapshot.bk_path, snapshot.bk_sheet, snapshot.vessel_voyage_raw,
                    snapshot.vessel_key, snapshot.voyage_key, snapshot.combined_key,
                    snapshot.container_count, snapshot.workbook_fingerprint,
                    snapshot.snapshot_hash, snapshot.invalid_container_count,
                    snapshot.duplicate_container_count, timestamp, group_id,
                ),
            )
            if cursor.rowcount != 1:
                raise ValueError("Nhóm đã phân bổ, đã đóng hoặc không tồn tại.")
            connection.execute(
                """
                UPDATE sea_freight_invoice_contributions
                SET vessel_voyage_raw = ?, vessel_name = ?, voyage_no = ?, updated_at = ?
                WHERE group_id = ? AND status = 'ACTIVE'
                """,
                (snapshot.vessel_voyage_raw, vessel_name, voyage_no, timestamp, group_id),
            )
            connection.execute(
                "DELETE FROM sea_freight_group_containers WHERE group_id = ?",
                (group_id,),
            )
            for order, record in enumerate(snapshot.containers):
                connection.execute(
                    """
                    INSERT INTO sea_freight_group_containers(
                        group_id, container, source_sheet, source_row,
                        source_sqt, departure_date, allocation_order
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (group_id, record.container, record.source_sheet, record.source_row,
                     record.source_sqt, record.departure_date, order),
                )
            self._recalculate(connection, group_id, timestamp)
            self._event(connection, group_id, "GROUP_MATCH_CHANGED", {
                "bk_sheet": snapshot.bk_sheet,
                "vessel_voyage_raw": snapshot.vessel_voyage_raw,
            }, timestamp)
        group = self.get_group(group_id)
        if group is None:
            raise KeyError("Nhóm không tồn tại hoặc đã đóng.")
        return group

    def add_contribution(self, group_id: int, payload: dict[str, Any]) -> InvoiceContribution:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            duplicate = connection.execute(
                """
                SELECT id FROM sea_freight_invoice_contributions
                WHERE fingerprint = ? AND status != 'REMOVED'
                """,
                (payload["fingerprint"],),
            ).fetchone()
            if duplicate is not None:
                raise DuplicateContributionError("Hóa đơn này đã được thêm vào đối soát.")
            cursor = connection.execute(
                """
                INSERT INTO sea_freight_invoice_contributions(
                    group_id, source_batch_id, source_item_index, source_sha256,
                    source_document_id, source_document_name,
                    invoice_no, invoice_date, bl, vessel_voyage_raw,
                    vessel_name, voyage_no, invoice_container_count,
                    container_count_basis, carrier, amount, fingerprint,
                    source_kind, status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'ACTIVE', ?, ?)
                """,
                (
                    group_id,
                    payload.get("source_batch_id"),
                    payload["source_item_index"],
                    payload["source_sha256"],
                    payload["source_document_id"],
                    payload["source_document_name"],
                    payload.get("invoice_no"),
                    payload.get("invoice_date"),
                    payload.get("bl"),
                    payload["vessel_voyage_raw"],
                    payload["vessel_name"],
                    payload["voyage_no"],
                    payload["invoice_container_count"],
                    payload["container_count_basis"],
                    payload.get("carrier"),
                    payload["amount"],
                    payload["fingerprint"],
                    payload.get("source_kind", "INITIAL"),
                    timestamp,
                    timestamp,
                ),
            )
            contribution_id = int(cursor.lastrowid)
            connection.execute(
                "UPDATE sea_freight_reconciliation_groups SET invoice_conflict_accepted = 0 WHERE id = ?",
                (group_id,),
            )
            if payload.get("source_batch_id") is not None:
                connection.execute(
                    """
                    UPDATE sea_freight_reconciliation_groups
                    SET primary_source_batch_id = COALESCE(primary_source_batch_id, ?),
                        primary_source_item_index = COALESCE(primary_source_item_index, ?)
                    WHERE id = ?
                    """,
                    (
                        int(payload["source_batch_id"]),
                        int(payload["source_item_index"]),
                        group_id,
                    ),
                )
            self._recalculate(connection, group_id, timestamp)
            self._event(connection, group_id, "CONTRIBUTION_ADDED", {
                "contribution_id": contribution_id,
                "invoice_no": payload.get("invoice_no"),
            }, timestamp)
        contribution = self.get_contribution(contribution_id)
        assert contribution is not None
        return contribution

    def add_contributions(
        self, group_id: int, payloads: list[dict[str, Any]]
    ) -> list[InvoiceContribution]:
        """Thêm nhiều hóa đơn nguyên tử; một dòng lỗi thì không dòng nào được lưu."""

        if not payloads:
            return []
        timestamp = _now()
        fingerprints = [str(payload["fingerprint"]) for payload in payloads]
        if len(fingerprints) != len(set(fingerprints)):
            raise DuplicateContributionError("Danh sách có hóa đơn bị trùng.")
        contribution_ids: list[int] = []
        with self.database.transaction(immediate=True) as connection:
            placeholders = ",".join("?" for _ in fingerprints)
            duplicate = connection.execute(
                f"SELECT id FROM sea_freight_invoice_contributions "
                f"WHERE fingerprint IN ({placeholders}) AND status != 'REMOVED' LIMIT 1",
                tuple(fingerprints),
            ).fetchone()
            if duplicate is not None:
                raise DuplicateContributionError(
                    "Có hóa đơn đã được thêm vào đối soát."
                )
            for payload in payloads:
                cursor = connection.execute(
                    """
                    INSERT INTO sea_freight_invoice_contributions(
                        group_id, source_batch_id, source_item_index, source_sha256,
                        source_document_id, source_document_name,
                        invoice_no, invoice_date, bl, vessel_voyage_raw,
                        vessel_name, voyage_no, invoice_container_count,
                        container_count_basis, carrier, amount, fingerprint,
                        source_kind, status, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'ACTIVE', ?, ?)
                    """,
                    (
                        group_id,
                        payload.get("source_batch_id"),
                        payload["source_item_index"],
                        payload["source_sha256"],
                        payload["source_document_id"],
                        payload["source_document_name"],
                        payload.get("invoice_no"),
                        payload.get("invoice_date"),
                        payload.get("bl"),
                        payload["vessel_voyage_raw"],
                        payload["vessel_name"],
                        payload["voyage_no"],
                        payload["invoice_container_count"],
                        payload["container_count_basis"],
                        payload.get("carrier"),
                        payload["amount"],
                        payload["fingerprint"],
                        payload.get("source_kind", "INITIAL"),
                        timestamp,
                        timestamp,
                    ),
                )
                contribution_ids.append(int(cursor.lastrowid))
                if payload.get("source_batch_id") is not None:
                    connection.execute(
                        """
                        UPDATE sea_freight_reconciliation_groups
                        SET primary_source_batch_id = COALESCE(primary_source_batch_id, ?),
                            primary_source_item_index = COALESCE(primary_source_item_index, ?)
                        WHERE id = ?
                        """,
                        (
                            int(payload["source_batch_id"]),
                            int(payload["source_item_index"]),
                            group_id,
                        ),
                    )
            connection.execute(
                "UPDATE sea_freight_reconciliation_groups "
                "SET invoice_conflict_accepted = 0 WHERE id = ?",
                (group_id,),
            )
            self._recalculate(connection, group_id, timestamp)
            self._event(
                connection,
                group_id,
                "CONTRIBUTIONS_ADDED",
                {
                    "contribution_ids": contribution_ids,
                    "invoice_nos": [payload.get("invoice_no") for payload in payloads],
                },
                timestamp,
            )
        result = [self.get_contribution(item_id) for item_id in contribution_ids]
        return [item for item in result if item is not None]

    def save_contributions(
        self,
        group_id: int,
        rows: list[dict[str, Any]],
        *,
        snapshot: BkContainerSnapshot | None = None,
        vessel_name: str | None = None,
        voyage_no: str | None = None,
        month: int | None = None,
        year: int | None = None,
        bk_issue: str | None = None,
    ) -> ReconciliationGroup:
        """Lưu toàn bộ bản nháp HĐ trong một transaction, giữ lịch sử dòng bị xóa."""

        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            group = connection.execute(
                "SELECT status FROM sea_freight_reconciliation_groups WHERE id = ?",
                (group_id,),
            ).fetchone()
            if group is None:
                raise KeyError("Không tìm thấy hồ sơ đối soát.")
            if str(group["status"]) in {"ALLOCATED", "POSTED", "CANCELLED"}:
                raise ValueError("Hồ sơ đã hoàn tất, không thể sửa.")
            if snapshot is not None:
                connection.execute(
                    """
                    UPDATE sea_freight_reconciliation_groups
                    SET bk_path = ?, bk_sheet = ?, vessel_voyage_raw = ?,
                        vessel_key = ?, voyage_key = ?, combined_key = ?,
                        bk_container_count = ?, bk_fingerprint = ?,
                        container_snapshot_hash = ?, reconciliation_month = ?,
                        reconciliation_year = ?, bk_issue = ?,
                        bk_invalid_container_count = ?, bk_duplicate_container_count = ?,
                        status = 'PENDING', updated_at = ?
                    WHERE id = ?
                    """,
                    (
                        snapshot.bk_path, snapshot.bk_sheet, snapshot.vessel_voyage_raw,
                        snapshot.vessel_key, snapshot.voyage_key, snapshot.combined_key,
                        snapshot.container_count, snapshot.workbook_fingerprint,
                        snapshot.snapshot_hash, month, year, bk_issue,
                        snapshot.invalid_container_count, snapshot.duplicate_container_count,
                        timestamp, group_id,
                    ),
                )
                connection.execute(
                    """
                    UPDATE sea_freight_invoice_contributions
                    SET vessel_voyage_raw = ?, vessel_name = ?, voyage_no = ?, updated_at = ?
                    WHERE group_id = ? AND status = 'ACTIVE'
                    """,
                    (snapshot.vessel_voyage_raw, vessel_name, voyage_no, timestamp, group_id),
                )
                connection.execute(
                    "DELETE FROM sea_freight_group_containers WHERE group_id = ?",
                    (group_id,),
                )
                for order, record in enumerate(snapshot.containers):
                    connection.execute(
                        """
                        INSERT INTO sea_freight_group_containers(
                            group_id, container, source_sheet, source_row,
                            source_sqt, departure_date, allocation_order
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            group_id, record.container, record.source_sheet,
                            record.source_row, record.source_sqt,
                            record.departure_date, order,
                        ),
                    )
            active_ids = {
                int(row["id"])
                for row in connection.execute(
                    "SELECT id FROM sea_freight_invoice_contributions WHERE group_id = ? AND status = 'ACTIVE'",
                    (group_id,),
                ).fetchall()
            }
            kept_ids: set[int] = set()
            for payload in rows:
                contribution_id = payload.get("id")
                if contribution_id is not None and int(contribution_id) in active_ids:
                    contribution_id = int(contribution_id)
                    kept_ids.add(contribution_id)
                    connection.execute(
                        """
                        UPDATE sea_freight_invoice_contributions
                        SET invoice_no = ?, invoice_date = ?, bl = ?,
                            invoice_container_count = ?, container_count_basis = ?,
                            carrier = ?, amount = ?, source_document_id = ?,
                            source_document_name = ?, updated_at = ?
                        WHERE id = ? AND group_id = ? AND status = 'ACTIVE'
                        """,
                        (
                            payload.get("invoice_no"), payload.get("invoice_date"),
                            payload.get("bl"), payload["invoice_container_count"],
                            payload.get("container_count_basis", "EXPLICIT"),
                            payload.get("carrier"), payload["amount"],
                            payload.get("source_document_id", "MANUAL"),
                            payload.get("source_document_name", "Dòng thêm thủ công"),
                            timestamp,
                            contribution_id, group_id,
                        ),
                    )
                    continue
                cursor = connection.execute(
                    """
                    INSERT INTO sea_freight_invoice_contributions(
                        group_id, source_batch_id, source_item_index, source_sha256,
                        source_document_id, source_document_name,
                        invoice_no, invoice_date, bl, vessel_voyage_raw,
                        vessel_name, voyage_no, invoice_container_count,
                        container_count_basis, carrier, amount, fingerprint,
                        source_kind, status, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'ACTIVE', ?, ?)
                    """,
                    (
                        group_id, payload.get("source_batch_id"),
                        int(payload.get("source_item_index", 0)),
                        str(payload.get("source_sha256") or f"MANUAL:{group_id}:{timestamp}"),
                        str(payload.get("source_document_id") or "MANUAL"),
                        str(payload.get("source_document_name") or "Dòng thêm thủ công"),
                        payload.get("invoice_no"), payload.get("invoice_date"),
                        payload.get("bl"), payload["vessel_voyage_raw"],
                        payload["vessel_name"], payload["voyage_no"],
                        payload["invoice_container_count"],
                        payload.get("container_count_basis", "EXPLICIT"),
                        payload.get("carrier"), payload["amount"], payload["fingerprint"],
                        payload.get("source_kind", "MANUAL"), timestamp, timestamp,
                    ),
                )
                kept_ids.add(int(cursor.lastrowid))
            removed_ids = active_ids - kept_ids
            if removed_ids:
                placeholders = ",".join("?" for _ in removed_ids)
                connection.execute(
                    f"UPDATE sea_freight_invoice_contributions SET status = 'REMOVED', removed_at = ?, updated_at = ? WHERE id IN ({placeholders})",
                    (timestamp, timestamp, *sorted(removed_ids)),
                )
            connection.execute(
                "UPDATE sea_freight_reconciliation_groups SET invoice_conflict_accepted = 0 WHERE id = ?",
                (group_id,),
            )
            self._recalculate(connection, group_id, timestamp)
            self._event(connection, group_id, "DRAFT_SAVED", {
                "row_count": len(rows),
                "snapshot_hash": snapshot.snapshot_hash if snapshot is not None else None,
            }, timestamp)
        result = self.get_group(group_id)
        assert result is not None
        return result

    def get_contribution(self, contribution_id: int) -> InvoiceContribution | None:
        row = self.database.query_one(
            "SELECT * FROM sea_freight_invoice_contributions WHERE id = ?",
            (contribution_id,),
        )
        return self._contribution(row) if row is not None else None

    def list_contributions(self, group_id: int) -> list[InvoiceContribution]:
        return [
            self._contribution(row)
            for row in self.database.query_all(
                """
                SELECT * FROM sea_freight_invoice_contributions
                WHERE group_id = ? AND status = 'ACTIVE' ORDER BY id
                """,
                (group_id,),
            )
        ]

    def list_contribution_history(self, group_id: int) -> list[InvoiceContribution]:
        """Trả cả HĐ đang dùng và HĐ đã gỡ để tra cứu lịch sử."""

        return [
            self._contribution(row)
            for row in self.database.query_all(
                """
                SELECT * FROM sea_freight_invoice_contributions
                WHERE group_id = ? ORDER BY id
                """,
                (group_id,),
            )
        ]

    def list_container_rows(self, group_id: int) -> list[dict[str, Any]]:
        return [
            dict(row)
            for row in self.database.query_all(
                """
                SELECT * FROM sea_freight_group_containers
                WHERE group_id = ? ORDER BY allocation_order, id
                """,
                (group_id,),
            )
        ]

    def managed_source_indices(self, batch_id: int) -> set[int]:
        return {
            int(row["source_item_index"])
            for row in self.database.query_all(
                """
                SELECT DISTINCT c.source_item_index
                FROM sea_freight_invoice_contributions AS c
                JOIN sea_freight_reconciliation_groups AS g ON g.id = c.group_id
                WHERE c.source_batch_id = ? AND c.status = 'ACTIVE'
                  AND g.status != 'CANCELLED'
                  AND g.is_current = 1
                """,
                (batch_id,),
            )
        }

    def groups_for_source_batch(self, batch_id: int) -> dict[int, ReconciliationGroup]:
        rows = self.database.query_all(
            """
            SELECT c.source_item_index, g.*
            FROM sea_freight_invoice_contributions AS c
            JOIN sea_freight_reconciliation_groups AS g ON g.id = c.group_id
            WHERE c.source_batch_id = ? AND c.status = 'ACTIVE'
              AND g.status != 'CANCELLED'
              AND g.is_current = 1
            ORDER BY g.revision_no, c.id
            """,
            (batch_id,),
        )
        return {
            int(row["source_item_index"]): self._group(row)
            for row in rows
        }

    def posting_groups_for_source_batch(self, batch_id: int) -> list[ReconciliationGroup]:
        """Các hồ sơ hiện hành mà batch nguồn sở hữu kết quả phân bổ."""

        return [
            self._group(row)
            for row in self.database.query_all(
                """
                SELECT * FROM sea_freight_reconciliation_groups
                WHERE primary_source_batch_id = ? AND is_current = 1
                  AND status != 'CANCELLED'
                ORDER BY primary_source_item_index, id
                """,
                (batch_id,),
            )
        ]

    def list_revisions(self, group_id: int) -> list[ReconciliationGroup]:
        current = self.get_group(group_id)
        if current is None:
            return []
        return [
            self._group(row)
            for row in self.database.query_all(
                """
                SELECT * FROM sea_freight_reconciliation_groups
                WHERE bk_path = ? AND bk_sheet = ? AND vessel_key = ? AND voyage_key = ?
                ORDER BY revision_no DESC, id DESC
                """,
                (current.bk_path, current.bk_sheet, current.vessel_key, current.voyage_key),
            )
        ]

    def create_revision(
        self,
        group_id: int,
        snapshot: BkContainerSnapshot,
    ) -> ReconciliationGroup:
        """Tạo phiên mới bất biến từ hồ sơ đã xác nhận/đã ghi BK."""

        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            old = connection.execute(
                "SELECT * FROM sea_freight_reconciliation_groups WHERE id = ?",
                (group_id,),
            ).fetchone()
            if old is None:
                raise KeyError("Không tìm thấy hồ sơ đối soát.")
            if not bool(old["is_current"]):
                raise ValueError("Chỉ có thể chạy lại phiên đối soát hiện hành.")
            if str(old["status"]) not in {"ALLOCATED", "POSTED"}:
                raise ValueError("Hồ sơ chưa hoàn tất nên không cần tạo phiên mới.")
            connection.execute(
                "UPDATE sea_freight_reconciliation_groups SET is_current = 0 WHERE id = ?",
                (group_id,),
            )
            if old["generated_batch_id"] is not None:
                connection.execute(
                    "UPDATE batches SET status = 'ARCHIVED' WHERE id = ? AND status != 'ARCHIVED'",
                    (int(old["generated_batch_id"]),),
                )
            cursor = connection.execute(
                """
                INSERT INTO sea_freight_reconciliation_groups(
                    bk_path, bk_sheet, vessel_voyage_raw, vessel_key, voyage_key,
                    combined_key, bk_container_count, received_container_count,
                    total_amount, status, bk_fingerprint, container_snapshot_hash,
                    output_carrier, invoice_conflict_accepted,
                    reconciliation_month, reconciliation_year, bk_issue,
                    bk_invalid_container_count, bk_duplicate_container_count,
                    revision_no, supersedes_group_id, is_current,
                    primary_source_batch_id, primary_source_item_index,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, 0, 0, 'PENDING', ?, ?, ?, 0,
                          ?, ?, NULL, ?, ?, ?, ?, 1, ?, ?, ?, ?)
                """,
                (
                    snapshot.bk_path, snapshot.bk_sheet, snapshot.vessel_voyage_raw,
                    snapshot.vessel_key, snapshot.voyage_key, snapshot.combined_key,
                    snapshot.container_count, snapshot.workbook_fingerprint,
                    snapshot.snapshot_hash, old["output_carrier"],
                    old["reconciliation_month"], old["reconciliation_year"],
                    snapshot.invalid_container_count, snapshot.duplicate_container_count,
                    int(old["revision_no"] or 1) + 1, group_id,
                    old["primary_source_batch_id"], old["primary_source_item_index"],
                    timestamp, timestamp,
                ),
            )
            new_id = int(cursor.lastrowid)
            connection.execute(
                """
                INSERT INTO sea_freight_invoice_contributions(
                    group_id, source_batch_id, source_item_index, source_sha256,
                    invoice_no, invoice_date, bl, vessel_voyage_raw, vessel_name,
                    voyage_no, invoice_container_count, container_count_basis,
                    carrier, amount, fingerprint, source_kind, status, created_at, updated_at
                )
                SELECT ?, source_batch_id, source_item_index, source_sha256,
                       invoice_no, invoice_date, bl, ?, vessel_name, voyage_no,
                       invoice_container_count, container_count_basis, carrier,
                       amount, fingerprint, source_kind, 'ACTIVE', ?, ?
                FROM sea_freight_invoice_contributions
                WHERE group_id = ? AND status = 'ACTIVE'
                ORDER BY id
                """,
                (new_id, snapshot.vessel_voyage_raw, timestamp, timestamp, group_id),
            )
            for order, record in enumerate(snapshot.containers):
                connection.execute(
                    """
                    INSERT INTO sea_freight_group_containers(
                        group_id, container, source_sheet, source_row,
                        source_sqt, departure_date, allocation_order
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (new_id, record.container, record.source_sheet, record.source_row,
                     record.source_sqt, record.departure_date, order),
                )
            self._recalculate(connection, new_id, timestamp)
            self._event(connection, group_id, "REVISION_SUPERSEDED", {"new_group_id": new_id}, timestamp)
            self._event(connection, new_id, "REVISION_CREATED", {"supersedes_group_id": group_id}, timestamp)
        created = self.get_group(new_id)
        assert created is not None
        return created

    def set_output_carrier(self, group_id: int, carrier: str) -> ReconciliationGroup:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            connection.execute(
                """
                UPDATE sea_freight_reconciliation_groups
                SET output_carrier = ?, updated_at = ? WHERE id = ?
                """,
                (carrier.strip(), timestamp, group_id),
            )
            self._recalculate(connection, group_id, timestamp)
            self._event(connection, group_id, "OUTPUT_CARRIER_SELECTED", {
                "carrier": carrier.strip()
            }, timestamp)
        group = self.get_group(group_id)
        assert group is not None
        return group

    def set_period_issue(
        self,
        group_id: int,
        *,
        month: int,
        year: int,
        bk_issue: str | None,
    ) -> ReconciliationGroup:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            connection.execute(
                """
                UPDATE sea_freight_reconciliation_groups
                SET reconciliation_month = ?, reconciliation_year = ?,
                    bk_issue = ?, updated_at = ?
                WHERE id = ?
                """,
                (month, year, bk_issue, timestamp, group_id),
            )
            self._recalculate(connection, group_id, timestamp)
        group = self.get_group(group_id)
        assert group is not None
        return group

    def accept_invoice_conflicts(self, group_id: int) -> ReconciliationGroup:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            connection.execute(
                "UPDATE sea_freight_reconciliation_groups SET invoice_conflict_accepted = 1, updated_at = ? WHERE id = ?",
                (timestamp, group_id),
            )
            self._recalculate(connection, group_id, timestamp)
            self._event(connection, group_id, "INVOICE_CONFLICT_ACCEPTED", {}, timestamp)
        group = self.get_group(group_id)
        assert group is not None
        return group

    def mark_allocated(self, group_id: int, generated_batch_id: int | None = None) -> None:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            cursor = connection.execute(
                """
                UPDATE sea_freight_reconciliation_groups
                SET status = 'ALLOCATED', generated_batch_id = ?, updated_at = ?
                WHERE id = ? AND status = 'READY'
                """,
                (generated_batch_id, timestamp, group_id),
            )
            if cursor.rowcount != 1:
                raise ValueError("Nhóm không còn ở trạng thái READY.")
            self._event(connection, group_id, "ALLOCATION_CONFIRMED", {
                "generated_batch_id": generated_batch_id
            }, timestamp)

    def mark_needs_recheck(self, group_id: int, reason: str) -> None:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            connection.execute(
                """
                UPDATE sea_freight_reconciliation_groups
                SET status = 'NEEDS_RECHECK', updated_at = ?
                WHERE id = ? AND status NOT IN ('POSTED','CANCELLED')
                """,
                (timestamp, group_id),
            )
            self._event(connection, group_id, "NEEDS_RECHECK", {"reason": reason}, timestamp)

    def mark_posted(self, group_id: int, run_id: int | None) -> None:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            cursor = connection.execute(
                """
                UPDATE sea_freight_reconciliation_groups
                SET status = 'POSTED', posted_run_id = ?, posted_at = ?, updated_at = ?
                WHERE id = ? AND status = 'ALLOCATED'
                """,
                (run_id, timestamp, timestamp, group_id),
            )
            if cursor.rowcount != 1:
                raise ValueError("Nhóm chưa được phân bổ hoặc đã được ghi trước đó.")
            self._event(connection, group_id, "POSTED", {"run_id": run_id}, timestamp)

    def remove_contribution(self, contribution_id: int) -> ReconciliationGroup:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            row = connection.execute(
                "SELECT group_id FROM sea_freight_invoice_contributions WHERE id = ? AND status = 'ACTIVE'",
                (contribution_id,),
            ).fetchone()
            if row is None:
                raise KeyError("Không tìm thấy hóa đơn đang hoạt động.")
            group_id = int(row["group_id"])
            connection.execute(
                "UPDATE sea_freight_invoice_contributions SET status = 'REMOVED', removed_at = ?, updated_at = ? WHERE id = ?",
                (timestamp, timestamp, contribution_id),
            )
            connection.execute(
                "UPDATE sea_freight_reconciliation_groups SET invoice_conflict_accepted = 0 WHERE id = ?",
                (group_id,),
            )
            self._recalculate(connection, group_id, timestamp)
            self._event(connection, group_id, "CONTRIBUTION_REMOVED", {"contribution_id": contribution_id}, timestamp)
        group = self.get_group(group_id)
        assert group is not None
        return group

    def cancel_group(self, group_id: int) -> None:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            cursor = connection.execute(
                "UPDATE sea_freight_reconciliation_groups SET status = 'CANCELLED', updated_at = ? WHERE id = ? AND status NOT IN ('POSTED','CANCELLED')",
                (timestamp, group_id),
            )
            if cursor.rowcount != 1:
                raise ValueError("Nhóm đã đóng hoặc không tồn tại.")
            connection.execute(
                """
                UPDATE sea_freight_invoice_contributions
                SET status = 'REMOVED', removed_at = ?, updated_at = ?
                WHERE group_id = ? AND status = 'ACTIVE'
                """,
                (timestamp, timestamp, group_id),
            )
            connection.execute(
                """
                UPDATE batches SET status = 'ARCHIVED'
                WHERE reconciliation_group_id = ?
                  AND status IN ('RECEIVED','REVIEWING','READY')
                """,
                (group_id,),
            )
            self._event(connection, group_id, "GROUP_CANCELLED", {}, timestamp)

    @staticmethod
    def _recalculate(connection: Any, group_id: int, timestamp: str) -> None:
        group = connection.execute(
            "SELECT * FROM sea_freight_reconciliation_groups WHERE id = ?",
            (group_id,),
        ).fetchone()
        totals = connection.execute(
            """
            SELECT COALESCE(SUM(invoice_container_count), 0) AS received,
                   COALESCE(SUM(amount), 0) AS total_amount,
                   COUNT(DISTINCT NULLIF(TRIM(carrier), '')) AS carrier_count
            FROM sea_freight_invoice_contributions
            WHERE group_id = ? AND status = 'ACTIVE'
            """,
            (group_id,),
        ).fetchone()
        invoice_conflict = connection.execute(
            """
            SELECT 1
            FROM sea_freight_invoice_contributions
            WHERE group_id = ? AND status = 'ACTIVE'
              AND NULLIF(TRIM(invoice_no), '') IS NOT NULL
            GROUP BY UPPER(TRIM(invoice_no))
            HAVING COUNT(DISTINCT fingerprint) > 1
            LIMIT 1
            """,
            (group_id,),
        ).fetchone()
        received = int(totals["received"])
        bk_count = int(group["bk_container_count"])
        if bk_count == 0:
            status = GroupStatus.BK_NOT_FOUND
        elif invoice_conflict is not None and not bool(group["invoice_conflict_accepted"]):
            status = GroupStatus.METADATA_CONFLICT
        elif received < bk_count:
            status = GroupStatus.PENDING
        elif received == bk_count:
            status = GroupStatus.READY
        else:
            status = GroupStatus.OVER_COUNT
        if group["status"] in {GroupStatus.ALLOCATED.value, GroupStatus.POSTED.value}:
            status = GroupStatus(str(group["status"]))
        connection.execute(
            """
            UPDATE sea_freight_reconciliation_groups
            SET received_container_count = ?, total_amount = ?, status = ?, updated_at = ?
            WHERE id = ?
            """,
            (received, int(totals["total_amount"]), status.value, timestamp, group_id),
        )

    @staticmethod
    def _event(
        connection: Any,
        group_id: int,
        event_type: str,
        payload: dict[str, Any],
        timestamp: str,
    ) -> None:
        connection.execute(
            """
            INSERT INTO sea_freight_reconciliation_events(
                group_id, event_type, payload_json, created_at
            ) VALUES (?, ?, ?, ?)
            """,
            (group_id, event_type, json.dumps(payload, ensure_ascii=False), timestamp),
        )

    @staticmethod
    def _group(row: Any) -> ReconciliationGroup:
        return ReconciliationGroup(
            id=int(row["id"]),
            bk_path=str(row["bk_path"]),
            bk_sheet=str(row["bk_sheet"]),
            vessel_voyage_raw=str(row["vessel_voyage_raw"]),
            vessel_key=str(row["vessel_key"]),
            voyage_key=str(row["voyage_key"]),
            combined_key=str(row["combined_key"]),
            bk_container_count=int(row["bk_container_count"]),
            received_container_count=int(row["received_container_count"]),
            total_amount=int(row["total_amount"]),
            status=GroupStatus(str(row["status"])),
            bk_fingerprint=row["bk_fingerprint"],
            container_snapshot_hash=row["container_snapshot_hash"],
            output_carrier=row["output_carrier"],
            invoice_conflict_accepted=bool(row["invoice_conflict_accepted"]),
            generated_batch_id=row["generated_batch_id"],
            posted_run_id=row["posted_run_id"],
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
            posted_at=row["posted_at"],
            reconciliation_month=row["reconciliation_month"],
            reconciliation_year=row["reconciliation_year"],
            bk_issue=row["bk_issue"],
            bk_invalid_container_count=int(row["bk_invalid_container_count"]),
            bk_duplicate_container_count=int(row["bk_duplicate_container_count"]),
            revision_no=int(row["revision_no"]),
            supersedes_group_id=row["supersedes_group_id"],
            is_current=bool(row["is_current"]),
            primary_source_batch_id=row["primary_source_batch_id"],
            primary_source_item_index=row["primary_source_item_index"],
        )

    @staticmethod
    def _contribution(row: Any) -> InvoiceContribution:
        return InvoiceContribution(**{key: row[key] for key in InvoiceContribution.__dataclass_fields__})
