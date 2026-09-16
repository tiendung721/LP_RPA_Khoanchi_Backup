from __future__ import annotations

import hashlib
import json
from datetime import datetime
from typing import Any

from app.database import Database

from .contracts import (
    BkContainerSnapshot,
    GroupStatus,
    InvoiceContribution,
    ReconciliationGroup,
    ReconciliationSource,
    VesselVoyageResolutionKind,
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

    def find_invoice_history(
        self,
        invoice_no: str,
        *,
        exclude_batch_id: int | None = None,
    ) -> list[tuple[InvoiceContribution, ReconciliationGroup]]:
        """Tìm HĐ đang hoạt động ở batch khác, ưu tiên phiên hiện hành mới nhất."""

        normalized = " ".join(str(invoice_no or "").split())
        if not normalized:
            return []
        parameters: list[Any] = [normalized]
        batch_filter = ""
        if exclude_batch_id is not None:
            batch_filter = "AND (c.source_batch_id IS NULL OR c.source_batch_id != ?)"
            parameters.append(int(exclude_batch_id))
        rows = self.database.query_all(
            f"""
            SELECT c.id AS contribution_id, g.id AS group_id
            FROM sea_freight_invoice_contributions AS c
            JOIN sea_freight_reconciliation_groups AS g ON g.id = c.group_id
            WHERE UPPER(TRIM(c.invoice_no)) = UPPER(TRIM(?))
              AND c.status = 'ACTIVE'
              AND g.status != 'CANCELLED'
              {batch_filter}
            ORDER BY g.is_current DESC, g.revision_no DESC, g.id DESC, c.id DESC
            """,
            tuple(parameters),
        )
        result: list[tuple[InvoiceContribution, ReconciliationGroup]] = []
        for row in rows:
            contribution = self.get_contribution(int(row["contribution_id"]))
            group = self.get_group(int(row["group_id"]))
            if contribution is not None and group is not None:
                result.append((contribution, group))
        return result

    def sync_duplicate_references(
        self,
        batch_id: int,
        desired: list[tuple[int, dict[str, Any]]],
    ) -> None:
        """Đồng bộ liên kết truy vết không tham gia tính toán cho một batch."""

        timestamp = _now()
        desired_by_index = {
            int(payload["source_item_index"]): (int(group_id), payload)
            for group_id, payload in desired
        }
        with self.database.transaction(immediate=True) as connection:
            existing_rows = connection.execute(
                """
                SELECT c.*
                FROM sea_freight_invoice_contributions AS c
                JOIN sea_freight_reconciliation_groups AS g ON g.id = c.group_id
                WHERE c.source_batch_id = ? AND c.status = 'DUPLICATE'
                  AND g.is_current = 1 AND g.status != 'CANCELLED'
                ORDER BY c.id
                """,
                (int(batch_id),),
            ).fetchall()
            retained: set[int] = set()
            for row in existing_rows:
                source_index = int(row["source_item_index"])
                target = desired_by_index.get(source_index)
                keep = (
                    target is not None
                    and int(row["group_id"]) == target[0]
                    and str(row["fingerprint"]) == str(target[1]["fingerprint"])
                    and source_index not in retained
                )
                if keep:
                    retained.add(source_index)
                    continue
                connection.execute(
                    """
                    UPDATE sea_freight_invoice_contributions
                    SET status = 'REMOVED', removed_at = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (timestamp, timestamp, int(row["id"])),
                )

            for source_index, (group_id, payload) in desired_by_index.items():
                if source_index in retained:
                    continue
                connection.execute(
                    """
                    INSERT INTO sea_freight_invoice_contributions(
                        group_id, source_batch_id, source_item_index, source_sha256,
                        source_document_id, source_document_name,
                        invoice_no, invoice_date, bl, vessel_voyage_raw,
                        vessel_name, voyage_no, invoice_container_count,
                        container_count_basis, carrier, amount, fingerprint,
                        source_kind, status, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                              'DUPLICATE', ?, ?)
                    """,
                    (
                        group_id,
                        int(batch_id),
                        source_index,
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
                        "DUPLICATE_REFERENCE",
                        timestamp,
                        timestamp,
                    ),
                )

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
                  AND status != 'CANCELLED'
                ORDER BY id DESC LIMIT 1
                """,
                (
                    snapshot.bk_path,
                    snapshot.bk_sheet,
                    snapshot.vessel_key,
                    snapshot.voyage_key,
                ),
            ).fetchone()
            if row is not None and str(row["status"]) in {"ALLOCATED", "POSTED"}:
                return self._group(row)
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

    def list_sources(self, group_id: int) -> list[ReconciliationSource]:
        return [
            self._source(row)
            for row in self.database.query_all(
                """
                SELECT * FROM sea_freight_reconciliation_sources
                WHERE group_id = ? ORDER BY source_order, id
                """,
                (group_id,),
            )
        ]

    def list_container_occurrences(self, group_id: int) -> list[dict[str, Any]]:
        return [
            dict(row)
            for row in self.database.query_all(
                """
                SELECT o.*, s.resolution_kind
                FROM sea_freight_container_occurrences AS o
                JOIN sea_freight_reconciliation_sources AS s ON s.id = o.source_id
                WHERE o.group_id = ?
                ORDER BY o.container, s.source_order, o.source_row, o.id
                """,
                (group_id,),
            )
        ]

    def unresolved_duplicate_containers(self, group_id: int) -> list[dict[str, Any]]:
        return [
            dict(row)
            for row in self.database.query_all(
                """
                SELECT container, COUNT(*) AS occurrence_count
                FROM sea_freight_container_occurrences
                WHERE group_id = ?
                GROUP BY container
                HAVING COUNT(*) > 1 AND SUM(selected) = 0
                ORDER BY container
                """,
                (group_id,),
            )
        ]

    def replace_group_sources(
        self,
        group_id: int,
        snapshots: list[BkContainerSnapshot],
    ) -> ReconciliationGroup:
        """Thay snapshot nhiều sheet và giữ lựa chọn cont trùng còn hợp lệ."""

        if not snapshots:
            raise ValueError("Hồ sơ phải có ít nhất một sheet đối soát.")
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            group = connection.execute(
                "SELECT status FROM sea_freight_reconciliation_groups WHERE id = ?",
                (group_id,),
            ).fetchone()
            if group is None:
                raise KeyError("Không tìm thấy hồ sơ đối soát.")
            had_persisted_sources = bool(
                connection.execute(
                    "SELECT 1 FROM sea_freight_reconciliation_sources "
                    "WHERE group_id = ? LIMIT 1",
                    (group_id,),
                ).fetchone()
            )
            old_selected = {
                str(row["container"]): (
                    str(row["source_sheet"]), int(row["source_row"])
                )
                for row in connection.execute(
                    "SELECT container, source_sheet, source_row "
                    "FROM sea_freight_container_occurrences "
                    "WHERE group_id = ? AND selected = 1 AND manually_selected = 1",
                    (group_id,),
                ).fetchall()
            } if had_persisted_sources else {}
            connection.execute(
                "DELETE FROM sea_freight_group_containers WHERE group_id = ?",
                (group_id,),
            )
            connection.execute(
                "DELETE FROM sea_freight_reconciliation_sources WHERE group_id = ?",
                (group_id,),
            )
            occurrences: dict[str, list[tuple[int, Any, int]]] = {}
            for order, snapshot in enumerate(snapshots):
                month = year = None
                text = snapshot.bk_sheet.strip()
                try:
                    month, year = int(text[1:3]), 2000 + int(text[4:6])
                except (ValueError, IndexError):
                    pass
                cursor = connection.execute(
                    """
                    INSERT INTO sea_freight_reconciliation_sources(
                        group_id, bk_sheet, reconciliation_month, reconciliation_year,
                        vessel_voyage_raw, vessel_name, voyage_no,
                        vessel_key, voyage_key, combined_key,
                        resolution_kind, container_count, invalid_container_count,
                        duplicate_container_count, snapshot_hash, source_order
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        group_id, snapshot.bk_sheet, month, year,
                        snapshot.vessel_voyage_raw,
                        snapshot.canonical_vessel_name or snapshot.vessel_key,
                        snapshot.canonical_voyage_no or snapshot.voyage_key,
                        snapshot.vessel_key,
                        snapshot.voyage_key, snapshot.combined_key,
                        snapshot.resolution_kind.value, snapshot.container_count,
                        snapshot.invalid_container_count,
                        snapshot.duplicate_container_count, snapshot.snapshot_hash, order,
                    ),
                )
                source_id = int(cursor.lastrowid)
                for record in snapshot.containers:
                    occurrences.setdefault(record.container, []).append(
                        (source_id, record, order)
                    )
            allocation_order = 0
            for container in sorted(occurrences):
                values = occurrences[container]
                old = old_selected.get(container)
                selected_index: int | None = None
                if len(values) == 1:
                    selected_index = 0
                elif old is not None:
                    selected_index = next(
                        (
                            index
                            for index, (_source_id, record, _order) in enumerate(values)
                            if (record.source_sheet, record.source_row) == old
                        ),
                        None,
                    )
                for index, (source_id, record, _order) in enumerate(values):
                    selected = int(index == selected_index)
                    connection.execute(
                        """
                        INSERT INTO sea_freight_container_occurrences(
                            group_id, source_id, container, source_sheet, source_row,
                            source_sqt, departure_date, selected, manually_selected
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            group_id, source_id, record.container,
                            record.source_sheet, record.source_row,
                            record.source_sqt, record.departure_date, selected,
                            int(selected and len(values) > 1),
                        ),
                    )
                    if selected:
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
                                record.departure_date, allocation_order,
                            ),
                        )
                        allocation_order += 1
            self._update_multi_source_aggregate(connection, group_id, timestamp)
            self._recalculate(connection, group_id, timestamp)
            self._event(
                connection,
                group_id,
                "BK_MULTI_SNAPSHOT",
                {"sheets": [snapshot.bk_sheet for snapshot in snapshots]},
                timestamp,
            )
        result = self.get_group(group_id)
        assert result is not None
        return result

    def select_container_occurrence(
        self,
        group_id: int,
        container: str,
        *,
        source_sheet: str,
        source_row: int,
    ) -> ReconciliationGroup:
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            occurrence = connection.execute(
                """
                SELECT * FROM sea_freight_container_occurrences
                WHERE group_id = ? AND container = ? AND source_sheet = ? AND source_row = ?
                """,
                (group_id, container, source_sheet, source_row),
            ).fetchone()
            if occurrence is None:
                raise ValueError("Vị trí container đã chọn không còn hợp lệ.")
            connection.execute(
                "UPDATE sea_freight_container_occurrences "
                "SET selected = 0, manually_selected = 0 "
                "WHERE group_id = ? AND container = ?",
                (group_id, container),
            )
            connection.execute(
                "UPDATE sea_freight_container_occurrences "
                "SET selected = 1, manually_selected = 1 WHERE id = ?",
                (int(occurrence["id"]),),
            )
            connection.execute(
                "DELETE FROM sea_freight_group_containers WHERE group_id = ? AND container = ?",
                (group_id, container),
            )
            next_order = int(
                connection.execute(
                    "SELECT COALESCE(MAX(allocation_order), -1) + 1 "
                    "FROM sea_freight_group_containers WHERE group_id = ?",
                    (group_id,),
                ).fetchone()[0]
            )
            connection.execute(
                """
                INSERT INTO sea_freight_group_containers(
                    group_id, container, source_sheet, source_row,
                    source_sqt, departure_date, allocation_order
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    group_id, occurrence["container"], occurrence["source_sheet"],
                    occurrence["source_row"], occurrence["source_sqt"],
                    occurrence["departure_date"], next_order,
                ),
            )
            self._update_multi_source_aggregate(connection, group_id, timestamp)
            self._recalculate(connection, group_id, timestamp)
        result = self.get_group(group_id)
        assert result is not None
        return result

    def managed_source_indices(self, batch_id: int) -> set[int]:
        return {
            int(row["source_item_index"])
            for row in self.database.query_all(
                """
                SELECT DISTINCT c.source_item_index
                FROM sea_freight_invoice_contributions AS c
                JOIN sea_freight_reconciliation_groups AS g ON g.id = c.group_id
                WHERE c.source_batch_id = ? AND c.status IN ('ACTIVE','DUPLICATE')
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
            WHERE c.source_batch_id = ? AND c.status IN ('ACTIVE','DUPLICATE')
              AND g.status != 'CANCELLED'
              AND g.is_current = 1
            ORDER BY g.revision_no,
                     CASE c.status WHEN 'DUPLICATE' THEN 0 ELSE 1 END,
                     c.id
            """,
            (batch_id,),
        )
        return {
            int(row["source_item_index"]): self._group(row)
            for row in rows
        }

    def current_source_links(
        self,
        batch_id: int,
        source_item_indices: set[int] | None = None,
    ) -> list[tuple[InvoiceContribution, ReconciliationGroup]]:
        """Các liên kết hiện hành của những dòng nguồn trong một batch."""

        parameters: list[Any] = [int(batch_id)]
        index_filter = ""
        if source_item_indices is not None:
            indices = sorted({int(value) for value in source_item_indices})
            if not indices:
                return []
            placeholders = ",".join("?" for _ in indices)
            index_filter = f"AND c.source_item_index IN ({placeholders})"
            parameters.extend(indices)
        rows = self.database.query_all(
            f"""
            SELECT c.id AS contribution_id, g.id AS group_id
            FROM sea_freight_invoice_contributions AS c
            JOIN sea_freight_reconciliation_groups AS g ON g.id = c.group_id
            WHERE c.source_batch_id = ?
              AND c.status IN ('ACTIVE','DUPLICATE')
              AND g.is_current = 1
              AND g.status != 'CANCELLED'
              {index_filter}
            ORDER BY g.id, c.id
            """,
            tuple(parameters),
        )
        result: list[tuple[InvoiceContribution, ReconciliationGroup]] = []
        for row in rows:
            contribution = self.get_contribution(int(row["contribution_id"]))
            group = self.get_group(int(row["group_id"]))
            if contribution is not None and group is not None:
                result.append((contribution, group))
        return result

    def remove_source_link(self, contribution_id: int) -> ReconciliationGroup:
        """Loại một HĐ/tham chiếu hiện hành và giữ lại lịch sử truy vết."""

        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            row = connection.execute(
                """
                SELECT group_id, status
                FROM sea_freight_invoice_contributions
                WHERE id = ? AND status IN ('ACTIVE','DUPLICATE')
                """,
                (int(contribution_id),),
            ).fetchone()
            if row is None:
                raise KeyError("Không tìm thấy liên kết hóa đơn đang hoạt động.")
            group_id = int(row["group_id"])
            status = str(row["status"])
            connection.execute(
                """
                UPDATE sea_freight_invoice_contributions
                SET status = 'REMOVED', removed_at = ?, updated_at = ?
                WHERE id = ?
                """,
                (timestamp, timestamp, int(contribution_id)),
            )
            if status == "ACTIVE":
                connection.execute(
                    """
                    UPDATE sea_freight_reconciliation_groups
                    SET invoice_conflict_accepted = 0
                    WHERE id = ?
                    """,
                    (group_id,),
                )
                self._recalculate(connection, group_id, timestamp)
            self._event(
                connection,
                group_id,
                "SOURCE_ROW_REMOVED",
                {"contribution_id": int(contribution_id), "link_status": status},
                timestamp,
            )
        group = self.get_group(group_id)
        assert group is not None
        return group

    def reindex_source_batch(
        self,
        batch_id: int,
        source_index_map: dict[int, int],
    ) -> None:
        """Cập nhật vị trí nguồn sau khi người dùng xóa dòng khỏi batch."""

        normalized = {
            int(old_index): int(new_index)
            for old_index, new_index in source_index_map.items()
        }
        timestamp = _now()
        with self.database.transaction(immediate=True) as connection:
            rows = connection.execute(
                """
                SELECT c.id, c.source_item_index
                FROM sea_freight_invoice_contributions AS c
                JOIN sea_freight_reconciliation_groups AS g ON g.id = c.group_id
                WHERE c.source_batch_id = ?
                  AND c.status IN ('ACTIVE','DUPLICATE')
                  AND g.is_current = 1
                  AND g.status != 'CANCELLED'
                """,
                (int(batch_id),),
            ).fetchall()
            for row in rows:
                old_index = int(row["source_item_index"])
                if old_index not in normalized:
                    continue
                connection.execute(
                    """
                    UPDATE sea_freight_invoice_contributions
                    SET source_item_index = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (normalized[old_index], timestamp, int(row["id"])),
                )

            owned_groups = connection.execute(
                """
                SELECT id
                FROM sea_freight_reconciliation_groups
                WHERE primary_source_batch_id = ? AND is_current = 1
                """,
                (int(batch_id),),
            ).fetchall()
            for row in owned_groups:
                group_id = int(row["id"])
                primary = connection.execute(
                    """
                    SELECT source_batch_id, source_item_index
                    FROM sea_freight_invoice_contributions
                    WHERE group_id = ?
                      AND status IN ('ACTIVE','DUPLICATE')
                      AND source_batch_id IS NOT NULL
                    ORDER BY CASE WHEN source_batch_id = ? THEN 0 ELSE 1 END,
                             source_item_index, id
                    LIMIT 1
                    """,
                    (group_id, int(batch_id)),
                ).fetchone()
                if primary is None:
                    continue
                connection.execute(
                    """
                    UPDATE sea_freight_reconciliation_groups
                    SET primary_source_batch_id = ?, primary_source_item_index = ?,
                        updated_at = ?
                    WHERE id = ?
                    """,
                    (
                        int(primary["source_batch_id"]),
                        int(primary["source_item_index"]),
                        timestamp,
                        group_id,
                    ),
                )

    def posting_groups_for_source_batch(self, batch_id: int) -> list[ReconciliationGroup]:
        """Các hồ sơ hiện hành được batch sở hữu hoặc liên kết để ghi BK."""

        return [
            self._group(row)
            for row in self.database.query_all(
                """
                SELECT * FROM sea_freight_reconciliation_groups AS g
                WHERE g.is_current = 1
                  AND (
                    g.primary_source_batch_id = ?
                    OR EXISTS (
                        SELECT 1
                        FROM sea_freight_invoice_contributions AS c
                        WHERE c.group_id = g.id AND c.source_batch_id = ?
                          AND c.status IN ('ACTIVE','DUPLICATE')
                    )
                  )
                  AND status != 'CANCELLED'
                ORDER BY primary_source_item_index, id
                """,
                (batch_id, batch_id),
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
                    source_document_id, source_document_name,
                    invoice_no, invoice_date, bl, vessel_voyage_raw, vessel_name,
                    voyage_no, invoice_container_count, container_count_basis,
                    carrier, amount, fingerprint, source_kind, status, created_at, updated_at
                )
                SELECT ?, source_batch_id, source_item_index, source_sha256,
                       source_document_id, source_document_name,
                       invoice_no, invoice_date, bl, ?, vessel_name, voyage_no,
                       invoice_container_count, container_count_basis, carrier,
                       amount, fingerprint, source_kind, status, ?, ?
                FROM sea_freight_invoice_contributions
                WHERE group_id = ? AND status IN ('ACTIVE','DUPLICATE')
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
                WHERE group_id = ? AND status != 'REMOVED'
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
        bk_issue = str(group["bk_issue"] or "")
        if bk_count == 0:
            status = GroupStatus.BK_NOT_FOUND
        elif bk_issue in {"SOURCE_UNRESOLVED", "DUPLICATE_UNRESOLVED"}:
            status = GroupStatus.METADATA_CONFLICT
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
    def _update_multi_source_aggregate(
        connection: Any,
        group_id: int,
        timestamp: str,
    ) -> None:
        selected_rows = connection.execute(
            """
            SELECT o.container
            FROM sea_freight_container_occurrences AS o
            JOIN sea_freight_reconciliation_sources AS s ON s.id = o.source_id
            WHERE o.group_id = ? AND o.selected = 1
            ORDER BY s.source_order, o.source_row, o.container
            """,
            (group_id,),
        ).fetchall()
        for order, row in enumerate(selected_rows):
            connection.execute(
                "UPDATE sea_freight_group_containers SET allocation_order = ? "
                "WHERE group_id = ? AND container = ?",
                (order, group_id, row["container"]),
            )
        sources = connection.execute(
            "SELECT * FROM sea_freight_reconciliation_sources "
            "WHERE group_id = ? ORDER BY source_order, id",
            (group_id,),
        ).fetchall()
        occurrences = connection.execute(
            "SELECT container, source_sheet, source_row, selected "
            "FROM sea_freight_container_occurrences WHERE group_id = ? "
            "ORDER BY container, source_sheet, source_row",
            (group_id,),
        ).fetchall()
        distinct_count = len({str(row["container"]) for row in occurrences})
        unresolved_duplicates = connection.execute(
            """
            SELECT COUNT(*) FROM (
                SELECT container
                FROM sea_freight_container_occurrences
                WHERE group_id = ?
                GROUP BY container
                HAVING COUNT(*) > 1 AND SUM(selected) = 0
            )
            """,
            (group_id,),
        ).fetchone()[0]
        unresolved_source = any(
            str(row["resolution_kind"]) in {"AMBIGUOUS", "NOT_FOUND"}
            for row in sources
        )
        issue = (
            "SOURCE_UNRESOLVED"
            if unresolved_source
            else "DUPLICATE_UNRESOLVED"
            if int(unresolved_duplicates)
            else None
        )
        payload = {
            "sources": [
                (row["bk_sheet"], row["combined_key"], row["snapshot_hash"])
                for row in sources
            ],
            "selected": [
                (row["container"], row["source_sheet"], row["source_row"])
                for row in occurrences
                if bool(row["selected"])
            ],
        }
        snapshot_hash = hashlib.sha256(
            json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        ).hexdigest()
        first = sources[0] if sources else None
        internal_duplicates = sum(int(row["duplicate_container_count"]) for row in sources)
        cross_duplicates = max(0, len(occurrences) - distinct_count)
        connection.execute(
            """
            UPDATE sea_freight_reconciliation_groups
            SET bk_container_count = ?, container_snapshot_hash = ?, bk_issue = ?,
                bk_invalid_container_count = ?, bk_duplicate_container_count = ?,
                bk_sheet = COALESCE(?, bk_sheet),
                reconciliation_month = COALESCE(?, reconciliation_month),
                reconciliation_year = COALESCE(?, reconciliation_year),
                vessel_voyage_raw = COALESCE(?, vessel_voyage_raw),
                vessel_key = COALESCE(?, vessel_key),
                voyage_key = COALESCE(?, voyage_key),
                combined_key = COALESCE(?, combined_key),
                updated_at = ?
            WHERE id = ?
            """,
            (
                distinct_count, snapshot_hash, issue,
                sum(int(row["invalid_container_count"]) for row in sources),
                internal_duplicates + cross_duplicates,
                first["bk_sheet"] if first is not None else None,
                first["reconciliation_month"] if first is not None else None,
                first["reconciliation_year"] if first is not None else None,
                first["vessel_voyage_raw"] if first is not None else None,
                first["vessel_key"] if first is not None else None,
                first["voyage_key"] if first is not None else None,
                first["combined_key"] if first is not None else None,
                timestamp, group_id,
            ),
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
    def _source(row: Any) -> ReconciliationSource:
        return ReconciliationSource(
            id=int(row["id"]),
            group_id=int(row["group_id"]),
            bk_sheet=str(row["bk_sheet"]),
            reconciliation_month=row["reconciliation_month"],
            reconciliation_year=row["reconciliation_year"],
            vessel_voyage_raw=str(row["vessel_voyage_raw"]),
            vessel_name=str(row["vessel_name"]),
            voyage_no=str(row["voyage_no"]),
            vessel_key=str(row["vessel_key"]),
            voyage_key=str(row["voyage_key"]),
            combined_key=str(row["combined_key"]),
            resolution_kind=VesselVoyageResolutionKind(str(row["resolution_kind"])),
            container_count=int(row["container_count"]),
            invalid_container_count=int(row["invalid_container_count"]),
            duplicate_container_count=int(row["duplicate_container_count"]),
            snapshot_hash=row["snapshot_hash"],
            source_order=int(row["source_order"]),
        )

    @staticmethod
    def _contribution(row: Any) -> InvoiceContribution:
        return InvoiceContribution(**{key: row[key] for key in InvoiceContribution.__dataclass_fields__})
