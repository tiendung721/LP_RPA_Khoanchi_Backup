from __future__ import annotations

import hashlib
import json
import logging
import sqlite3
from pathlib import Path
from typing import Iterable

from app.models import DataRow

from .contracts import (
    BkContainerSnapshot,
    GroupStatus,
    InvoiceHistoryMatch,
    InvoiceHistoryMatchKind,
    ReconciliationGroup,
    ReconciliationOpenResult,
    ReconciliationSourceSelection,
    SupplementImportResult,
    VesselVoyageResolutionKind,
    VesselVoyageSuggestion,
)
from .container_numbers import allocate_integer_amount
from .matching import (
    BkVesselMatcher,
    SeaFreightMatchError,
    normalize_match_key,
    vessel_voyage_alias_equivalent,
    vessel_voyage_keys,
    vessel_voyage_text,
)
from .repository import DuplicateContributionError, SeaFreightRepository


LOGGER = logging.getLogger(__name__)


class SeaFreightReconciliationError(RuntimeError):
    pass


class VesselVoyageNotFoundError(SeaFreightReconciliationError):
    def __init__(
        self,
        *,
        vessel_voyage: str,
        bk_sheet: str,
        suggestions: tuple[VesselVoyageSuggestion, ...] = (),
        sheet_missing: bool = False,
        ambiguous: bool = False,
    ) -> None:
        self.vessel_voyage = vessel_voyage
        self.bk_sheet = bk_sheet
        self.suggestions = suggestions
        self.sheet_missing = sheet_missing
        self.ambiguous = ambiguous
        if sheet_missing:
            message = f"Không tìm thấy sheet BK {bk_sheet}."
        else:
            message = (
                f'Không tìm thấy tàu/chuyến “{vessel_voyage}” trong sheet {bk_sheet}. '
                "Hãy kiểm tra và sửa lại thông tin tàu/chuyến, sau đó đối soát lại."
            )
        super().__init__(message)


class SeaFreightReconciliationService:
    def __init__(
        self,
        repository: SeaFreightRepository,
        *,
        matcher: BkVesselMatcher | None = None,
    ) -> None:
        self.repository = repository
        self.matcher = matcher or BkVesselMatcher()

    @staticmethod
    def period_sheet(month: int, year: int) -> str:
        if not 1 <= month <= 12 or not 2000 <= year <= 2999:
            raise ValueError("Tháng/năm đối soát không hợp lệ.")
        return f"T{month:02d} {year % 100:02d}"

    @staticmethod
    def period_from_sheet(sheet: str) -> tuple[int | None, int | None]:
        text = str(sheet).strip()
        try:
            return int(text[1:3]), 2000 + int(text[4:6])
        except (TypeError, ValueError, IndexError):
            return None, None

    def sheet_names(self, bk_path: str | Path) -> tuple[str, ...]:
        return self.matcher.sheet_names(bk_path)

    def inspect(
        self, row: DataRow, *, bk_path: str | Path, bk_sheet: str
    ) -> BkContainerSnapshot:
        self._require_candidate(row)
        return self.matcher.snapshot(
            bk_path,
            bk_sheet,
            vessel_voyage_raw=str(row.vessel_voyage_raw),
            vessel_name=str(row.vessel_name),
            voyage_no=str(row.voyage_no),
        )

    def resolve_for_period(
        self,
        row: DataRow,
        *,
        bk_path: str | Path,
        month: int,
        year: int,
    ) -> BkContainerSnapshot:
        """Xem trước canonical identity để UI gom đúng các hóa đơn tương đương."""

        snapshot, _issue = self._snapshot_for_period(
            row,
            bk_path=bk_path,
            month=month,
            year=year,
        )
        return snapshot

    def inspect_sources(
        self,
        row: DataRow,
        *,
        bk_path: str | Path,
        selections: Iterable[ReconciliationSourceSelection],
    ) -> tuple[BkContainerSnapshot, ...]:
        """Đọc độc lập từng sheet, giữ canonical identity của chính sheet đó."""

        result: list[BkContainerSnapshot] = []
        for selection in selections:
            probe = row.copy_with(
                vessel_name=selection.vessel_name,
                voyage_no=selection.voyage_no,
            )
            result.append(
                self.inspect(probe, bk_path=bk_path, bk_sheet=selection.bk_sheet)
            )
        return tuple(result)

    def _snapshot_for_period(
        self,
        row: DataRow,
        *,
        bk_path: str | Path,
        month: int,
        year: int,
    ) -> tuple[BkContainerSnapshot, str | None]:
        self._require_candidate(row)
        sheet = self.period_sheet(month, year)
        try:
            names = self.sheet_names(bk_path)
        except Exception:
            names = ()
        if sheet not in names:
            vessel_key, voyage_key, combined_key = vessel_voyage_keys(
                row.vessel_name, row.voyage_no
            )
            return (
                BkContainerSnapshot(
                    bk_path=str(Path(bk_path).expanduser().resolve()),
                    bk_sheet=sheet,
                    vessel_voyage_raw=vessel_voyage_text(
                        row.vessel_name, row.voyage_no
                    ),
                    vessel_key=vessel_key,
                    voyage_key=voyage_key,
                    combined_key=combined_key,
                    workbook_fingerprint="",
                    snapshot_hash="",
                    containers=(),
                ),
                "SHEET_MISSING",
            )
        snapshot = self.inspect(row, bk_path=bk_path, bk_sheet=sheet)
        return snapshot, "VESSEL_NOT_FOUND" if snapshot.container_count == 0 else None

    def open_or_create(
        self,
        row: DataRow,
        *,
        bk_path: str | Path,
        month: int,
        year: int,
        source_batch_id: int | None,
        source_item_index: int,
        source_sha256: str,
    ) -> ReconciliationGroup:
        return self.open_or_create_many(
            [(source_item_index, row)],
            bk_path=bk_path,
            month=month,
            year=year,
            source_batch_id=source_batch_id,
            source_sha256=source_sha256,
        ).group

    def attach(
        self,
        row: DataRow,
        *,
        bk_path: str | Path,
        bk_sheet: str,
        source_batch_id: int | None,
        source_item_index: int,
        source_sha256: str,
        source_kind: str = "INITIAL",
    ) -> ReconciliationGroup:
        snapshot = self.inspect(row, bk_path=bk_path, bk_sheet=bk_sheet)
        self._require_snapshot_match(row, snapshot)
        month, year = self.period_from_sheet(bk_sheet)
        existing = self.repository.find_open_group(
            snapshot.bk_path, snapshot.bk_sheet, snapshot.vessel_key, snapshot.voyage_key
        )
        if existing is None:
            latest = self.repository.find_latest_group(
                snapshot.bk_path, snapshot.bk_sheet, snapshot.vessel_key, snapshot.voyage_key
            )
            if latest is not None and latest.status is GroupStatus.POSTED:
                raise SeaFreightReconciliationError("Hồ sơ đã hoàn tất, không thể thêm HĐ.")
        if existing is not None and existing.status in {GroupStatus.ALLOCATED, GroupStatus.POSTED}:
            raise SeaFreightReconciliationError("Hồ sơ đã hoàn tất, không thể thêm HĐ.")
        if existing is not None:
            invoice_key = self._invoice_key(row.invoice_no)
            if invoice_key and invoice_key in {
                self._invoice_key(item.invoice_no)
                for item in self.repository.list_contributions(existing.id)
            }:
                raise SeaFreightReconciliationError("HĐ đã có trong hồ sơ đối soát.")
        group = self.repository.upsert_snapshot(
            snapshot,
            month=month,
            year=year,
            bk_issue=None,
        )
        group = self.repository.replace_group_sources(group.id, [snapshot])
        self._add_row(
            group.id,
            row.copy_with(
                vessel_name=(snapshot.canonical_vessel_name or row.vessel_name),
                voyage_no=(snapshot.canonical_voyage_no or row.voyage_no),
            ),
            source_batch_id=source_batch_id,
            source_item_index=source_item_index,
            source_sha256=source_sha256,
            source_kind=source_kind,
        )
        result = self.repository.get_group(group.id)
        assert result is not None
        return result

    def open_or_create_many(
        self,
        rows: Iterable[tuple[int, DataRow]],
        *,
        bk_path: str | Path,
        month: int | None = None,
        year: int | None = None,
        source_selections: Iterable[ReconciliationSourceSelection] | None = None,
        source_batch_id: int | None,
        source_sha256: str,
    ) -> ReconciliationOpenResult:
        """Gom các HĐ cùng tàu/chuyến vào một hồ sơ bằng một lần ghi nguyên tử."""

        selected = list(rows)
        if not selected:
            raise SeaFreightReconciliationError("Chưa chọn hóa đơn cước biển.")
        _, first_row = selected[0]
        selections = list(source_selections or ())
        if selections:
            selections.sort(
                key=lambda item: (
                    self.period_from_sheet(item.bk_sheet)[1] or 9999,
                    self.period_from_sheet(item.bk_sheet)[0] or 99,
                    item.bk_sheet,
                )
            )
            snapshots = list(
                self.inspect_sources(
                    first_row, bk_path=bk_path, selections=selections
                )
            )
            snapshot = next(
                (item for item in snapshots if item.container_count > 0), snapshots[0]
            )
            issue = (
                "SOURCE_UNRESOLVED"
                if any(item.container_count == 0 for item in snapshots)
                else None
            )
            primary_month, primary_year = self.period_from_sheet(snapshots[0].bk_sheet)
        else:
            if month is None or year is None:
                raise SeaFreightReconciliationError("Chưa chọn sheet đối soát.")
            snapshot, issue = self._snapshot_for_period(
                first_row, bk_path=bk_path, month=month, year=year
            )
            self._require_snapshot_match(first_row, snapshot, issue=issue)
            snapshots = [snapshot]
            primary_month, primary_year = month, year
        canonical_vessel_name = (
            snapshot.canonical_vessel_name or str(first_row.vessel_name or "").strip()
        )
        canonical_voyage_no = (
            snapshot.canonical_voyage_no or str(first_row.voyage_no or "").strip()
        )
        LOGGER.info(
            "Kết quả resolve tàu/chuyến: %s -> %s",
            snapshot.resolution_kind.value,
            snapshot.vessel_voyage_raw,
        )
        for _, row in selected:
            self._require_candidate(row)
            if not vessel_voyage_alias_equivalent(
                row.vessel_name,
                row.voyage_no,
                canonical_vessel_name,
                canonical_voyage_no,
            ):
                raise SeaFreightReconciliationError(
                    "Các hóa đơn được chọn không cùng tàu/chuyến."
                )
        current = self.repository.find_latest_group(
            snapshot.bk_path,
            snapshot.bk_sheet,
            snapshot.vessel_key,
            snapshot.voyage_key,
        )
        if current is not None and current.status in {
            GroupStatus.ALLOCATED,
            GroupStatus.POSTED,
        }:
            contributions = self.repository.list_contributions(current.id)
            duplicate_indices = tuple(
                source_index
                for source_index, row in selected
                if any(
                    self._invoice_key(row.invoice_no)
                    and self._invoice_key(row.invoice_no)
                    == self._invoice_key(item.invoice_no)
                    and not self._history_differences(row, item)
                    for item in contributions
                )
            )
            return ReconciliationOpenResult(
                current,
                duplicate_source_indices=duplicate_indices,
                requires_revision=len(duplicate_indices) != len(selected),
                resolution_kind=snapshot.resolution_kind,
                canonical_vessel_name=canonical_vessel_name,
                canonical_voyage_no=canonical_voyage_no,
            )

        try:
            group = self.repository.upsert_snapshot(
                snapshots[0], month=primary_month, year=primary_year, bk_issue=issue
            )
        except sqlite3.IntegrityError as exc:
            # Một thao tác đồng thời có thể tạo hồ sơ sau bước tra cứu. Luôn đổi
            # va chạm DB thành kết quả/ngữ nghĩa nghiệp vụ thay vì lộ lỗi SQLite.
            current = self.repository.find_latest_group(
                snapshot.bk_path,
                snapshot.bk_sheet,
                snapshot.vessel_key,
                snapshot.voyage_key,
            )
            if current is None:
                raise SeaFreightReconciliationError(
                    "Không thể mở hồ sơ đối soát hiện hành."
                ) from exc
            if current.status in {GroupStatus.ALLOCATED, GroupStatus.POSTED}:
                return ReconciliationOpenResult(
                    current,
                    requires_revision=True,
                    resolution_kind=snapshot.resolution_kind,
                    canonical_vessel_name=canonical_vessel_name,
                    canonical_voyage_no=canonical_voyage_no,
                )
            group = current
        group = self.repository.replace_group_sources(group.id, snapshots)
        existing_invoices = {
            self._invoice_key(item.invoice_no)
            for item in self.repository.list_contributions(group.id)
            if self._invoice_key(item.invoice_no)
        }
        payloads: list[dict[str, object]] = []
        added_indices: list[int] = []
        duplicate_indices: list[int] = []
        for source_item_index, row in selected:
            invoice_key = self._invoice_key(row.invoice_no)
            if invoice_key and invoice_key in existing_invoices:
                duplicate_indices.append(source_item_index)
                continue
            payloads.append(
                self._row_payload(
                    row.copy_with(
                        vessel_name=canonical_vessel_name,
                        voyage_no=canonical_voyage_no,
                    ),
                    source_batch_id=source_batch_id,
                    source_item_index=source_item_index,
                    source_sha256=source_sha256,
                    source_kind="INITIAL",
                )
            )
            added_indices.append(source_item_index)
            if invoice_key:
                existing_invoices.add(invoice_key)
        if payloads:
            self.repository.add_contributions(group.id, payloads)
        refreshed = self.repository.get_group(group.id)
        assert refreshed is not None
        return ReconciliationOpenResult(
            refreshed,
            added_source_indices=tuple(added_indices),
            duplicate_source_indices=tuple(duplicate_indices),
            resolution_kind=snapshot.resolution_kind,
            canonical_vessel_name=canonical_vessel_name,
            canonical_voyage_no=canonical_voyage_no,
        )

    def inspect_invoice_history(
        self,
        row: DataRow,
        *,
        source_batch_id: int | None,
    ) -> InvoiceHistoryMatch:
        invoice_key = self._invoice_key(row.invoice_no)
        if not invoice_key:
            return InvoiceHistoryMatch(InvoiceHistoryMatchKind.NONE)
        first_conflict: InvoiceHistoryMatch | None = None
        history = self.repository.find_invoice_history(
            str(row.invoice_no or ""),
            exclude_batch_id=source_batch_id,
        )
        current_history = [item for item in history if item[1].is_current]
        for contribution, group in current_history or history:
            differing = self._history_differences(row, contribution)
            if not differing:
                return InvoiceHistoryMatch(
                    InvoiceHistoryMatchKind.EXACT,
                    group,
                    contribution,
                )
            if first_conflict is None:
                first_conflict = InvoiceHistoryMatch(
                    InvoiceHistoryMatchKind.CONFLICT,
                    group,
                    contribution,
                    differing,
                )
        return first_conflict or InvoiceHistoryMatch(InvoiceHistoryMatchKind.NONE)

    def sync_batch_history(
        self,
        rows: Iterable[tuple[int, DataRow]],
        *,
        source_batch_id: int,
        source_sha256: str,
    ) -> dict[int, InvoiceHistoryMatch]:
        """Phân loại lịch sử và đồng bộ link DUPLICATE của các dòng hiện tại."""

        matches: dict[int, InvoiceHistoryMatch] = {}
        desired: list[tuple[int, dict[str, object]]] = []
        for source_item_index, row in rows:
            if row.fee != "CB" or row.cont not in (None, ""):
                continue
            match = self.inspect_invoice_history(
                row,
                source_batch_id=source_batch_id,
            )
            matches[source_item_index] = match
            if match.kind is not InvoiceHistoryMatchKind.EXACT or match.group is None:
                continue
            desired.append(
                (
                    match.group.id,
                    self._row_payload(
                        row,
                        source_batch_id=source_batch_id,
                        source_item_index=source_item_index,
                        source_sha256=source_sha256,
                        source_kind="DUPLICATE_REFERENCE",
                    ),
                )
            )
        self.repository.sync_duplicate_references(source_batch_id, desired)
        return matches

    def apply_batch_row_deletions(
        self,
        *,
        source_batch_id: int,
        deleted_source_indices: set[int],
        remaining_source_indices: dict[int, int],
    ) -> None:
        """Đồng bộ hồ sơ sau khi các dòng đã được lưu xóa khỏi batch nguồn."""

        deleted = {int(value) for value in deleted_source_indices}
        if not deleted:
            return

        links = self.repository.current_source_links(source_batch_id, deleted)
        completed_group_ids = {
            group.id
            for contribution, group in links
            if contribution.status == "ACTIVE"
            and group.status in {GroupStatus.ALLOCATED, GroupStatus.POSTED}
        }
        for group_id in sorted(completed_group_ids):
            # Phiên đã xác nhận/đã ghi BK là bất biến. Tạo phiên hiện hành mới
            # trước khi loại HĐ để lịch sử cũ vẫn được giữ nguyên.
            self.create_revision(group_id)

        # Sau khi tạo revision, ID contribution đã thay đổi nên phải đọc lại.
        links = self.repository.current_source_links(source_batch_id, deleted)
        affected_group_ids: set[int] = set()
        for contribution, group in links:
            self.repository.remove_source_link(contribution.id)
            affected_group_ids.add(group.id)

        for group_id in sorted(affected_group_ids):
            group = self.repository.get_group(group_id)
            if (
                group is not None
                and group.is_current
                and group.status is not GroupStatus.CANCELLED
                and not self.repository.list_contributions(group_id)
            ):
                self.repository.cancel_group(group_id)

        self.repository.reindex_source_batch(
            source_batch_id,
            remaining_source_indices,
        )

    @staticmethod
    def _history_differences(
        row: DataRow,
        contribution: object,
    ) -> tuple[str, ...]:
        text = lambda value: " ".join(str(value or "").split()).casefold()
        comparisons = (
            ("ngày HĐ", text(row.invoice_date), text(getattr(contribution, "invoice_date"))),
            ("B/L", text(row.bl), text(getattr(contribution, "bl"))),
            (
                "tàu/chuyến",
                True,
                vessel_voyage_alias_equivalent(
                    row.vessel_name,
                    row.voyage_no,
                    getattr(contribution, "vessel_name"),
                    getattr(contribution, "voyage_no"),
                ),
            ),
            (
                "số cont",
                row.invoice_container_count,
                getattr(contribution, "invoice_container_count"),
            ),
            ("số tiền", row.amount, getattr(contribution, "amount")),
            ("bên vận tải", text(row.carrier), text(getattr(contribution, "carrier"))),
        )
        return tuple(label for label, current, old in comparisons if current != old)

    def _add_row(
        self,
        group_id: int,
        row: DataRow,
        *,
        source_batch_id: int | None,
        source_item_index: int,
        source_sha256: str,
        source_kind: str,
    ) -> None:
        self.repository.add_contribution(
            group_id,
            self._row_payload(
                row,
                source_batch_id=source_batch_id,
                source_item_index=source_item_index,
                source_sha256=source_sha256,
                source_kind=source_kind,
            ),
        )

    def _row_payload(
        self,
        row: DataRow,
        *,
        source_batch_id: int | None,
        source_item_index: int,
        source_sha256: str,
        source_kind: str,
    ) -> dict[str, object]:
        return {
                "source_batch_id": source_batch_id,
                "source_item_index": source_item_index,
                "source_sha256": source_sha256,
                "source_document_id": row.source_document_id,
                "source_document_name": row.source_document_name,
                "invoice_no": row.invoice_no,
                "invoice_date": row.invoice_date,
                "bl": row.bl,
                "vessel_voyage_raw": row.vessel_voyage_raw,
                "vessel_name": row.vessel_name,
                "voyage_no": row.voyage_no,
                "invoice_container_count": row.invoice_container_count,
                "container_count_basis": row.container_count_basis,
                "carrier": row.carrier,
                "amount": row.amount,
                "fingerprint": self.contribution_fingerprint(row, source_sha256=source_sha256),
                "source_kind": source_kind,
            }

    def refresh_group(self, group_id: int) -> ReconciliationGroup:
        group = self._require_group(group_id)
        contributions = self.repository.list_contributions(group_id)
        if not contributions:
            raise SeaFreightReconciliationError("Hồ sơ chưa có HĐ.")
        first = contributions[0]
        row = self._contribution_row(first)
        sources = self.repository.list_sources(group_id)
        if sources:
            snapshots = list(
                self.inspect_sources(
                    row,
                    bk_path=group.bk_path,
                    selections=[
                        ReconciliationSourceSelection(
                            item.bk_sheet, item.vessel_name, item.voyage_no
                        )
                        for item in sources
                    ],
                )
            )
            return self.repository.replace_group_sources(group_id, snapshots)
        month = group.reconciliation_month
        year = group.reconciliation_year
        if month is None or year is None:
            month, year = self.period_from_sheet(group.bk_sheet)
        if month is None or year is None:
            raise SeaFreightReconciliationError("Chưa xác định được tháng đối soát.")
        snapshot, issue = self._snapshot_for_period(
            row, bk_path=group.bk_path, month=month, year=year
        )
        return self.repository.upsert_snapshot(
            snapshot, month=month, year=year, bk_issue=issue
        )

    def update_sources(
        self,
        group_id: int,
        selections: Iterable[ReconciliationSourceSelection],
    ) -> ReconciliationGroup:
        group = self._require_group(group_id)
        if not group.is_current or group.status in {
            GroupStatus.ALLOCATED, GroupStatus.POSTED, GroupStatus.CANCELLED
        }:
            raise SeaFreightReconciliationError("Hồ sơ đã hoàn tất, không thể sửa nguồn.")
        contributions = self.repository.list_contributions(group_id)
        if not contributions:
            raise SeaFreightReconciliationError("Hồ sơ chưa có HĐ.")
        selected = list(selections)
        if not selected:
            raise SeaFreightReconciliationError("Hồ sơ phải có ít nhất một sheet đối soát.")
        selected.sort(
            key=lambda item: (
                self.period_from_sheet(item.bk_sheet)[1] or 9999,
                self.period_from_sheet(item.bk_sheet)[0] or 99,
                item.bk_sheet,
            )
        )
        snapshots = list(
            self.inspect_sources(
                self._contribution_row(contributions[0]),
                bk_path=group.bk_path,
                selections=selected,
            )
        )
        return self.repository.replace_group_sources(group_id, snapshots)

    def select_container_source(
        self,
        group_id: int,
        container: str,
        *,
        source_sheet: str,
        source_row: int,
    ) -> ReconciliationGroup:
        return self.repository.select_container_occurrence(
            group_id,
            container,
            source_sheet=source_sheet,
            source_row=source_row,
        )

    def save_group(
        self,
        group_id: int,
        *,
        vessel_voyage_raw: str,
        vessel_name: str,
        voyage_no: str,
        invoices: list[dict[str, object]],
    ) -> ReconciliationGroup:
        group = self._require_group(group_id)
        if not group.is_current or group.status in {
            GroupStatus.ALLOCATED, GroupStatus.POSTED, GroupStatus.CANCELLED
        }:
            raise SeaFreightReconciliationError("Hồ sơ đã hoàn tất, không thể sửa.")
        vessel_voyage_keys(vessel_name, voyage_no)
        effective_vessel_voyage = vessel_voyage_text(vessel_name, voyage_no)
        seen_invoices: set[str] = set()
        normalized: list[dict[str, object]] = []
        for index, item in enumerate(invoices):
            invoice_no = str(item.get("invoice_no") or "").strip()
            if not invoice_no:
                raise SeaFreightReconciliationError(f"Dòng HĐ {index + 1} chưa có số HĐ.")
            key = self._invoice_key(invoice_no)
            if key in seen_invoices:
                raise SeaFreightReconciliationError(f"Số HĐ {invoice_no} đang bị trùng.")
            seen_invoices.add(key)
            try:
                count = int(item.get("invoice_container_count"))
                amount = int(item.get("amount"))
            except (TypeError, ValueError) as exc:
                raise SeaFreightReconciliationError(
                    f"Dòng HĐ {index + 1} có số cont hoặc số tiền không hợp lệ."
                ) from exc
            if count <= 0 or amount < 0:
                raise SeaFreightReconciliationError(
                    f"Dòng HĐ {index + 1}: số cont phải lớn hơn 0 và số tiền không được âm."
                )
            invoice_date = str(item.get("invoice_date") or "").strip() or None
            if invoice_date:
                from datetime import date

                try:
                    if date.fromisoformat(invoice_date).isoformat() != invoice_date:
                        raise ValueError
                except ValueError as exc:
                    raise SeaFreightReconciliationError(
                        f"Ngày HĐ dòng {index + 1} phải có dạng YYYY-MM-DD."
                    ) from exc
            source_sha = str(item.get("source_sha256") or f"MANUAL:{group_id}:{index}:{invoice_no}")
            row = DataRow(
                cont=None,
                bl=str(item.get("bl") or "").strip() or None,
                fee="CB",
                rule="HD",
                amount=amount,
                invoice_no=invoice_no,
                carrier=str(item.get("carrier") or "").strip() or None,
                vessel_voyage_raw=effective_vessel_voyage,
                vessel_name=vessel_name.strip(),
                voyage_no=voyage_no.strip(),
                invoice_container_count=count,
                container_count_basis="EXPLICIT",
                invoice_date=invoice_date,
                source_document_id=str(item.get("source_document_id") or "MANUAL"),
                source_document_name=str(
                    item.get("source_document_name") or "Dòng thêm thủ công"
                ),
            )
            normalized.append(
                {
                    **item,
                    **row.to_object(),
                    "id": item.get("id"),
                    "source_sha256": source_sha,
                    "fingerprint": self.contribution_fingerprint(row, source_sha256=source_sha),
                }
            )
        month = group.reconciliation_month
        year = group.reconciliation_year
        if month is None or year is None:
            month, year = self.period_from_sheet(group.bk_sheet)
        if month is None or year is None:
            raise SeaFreightReconciliationError("Chưa xác định được tháng đối soát.")
        probe = DataRow(
            cont=None, bl=None, fee="CB", rule="HD", amount=0,
            invoice_no="PROBE", vessel_voyage_raw=effective_vessel_voyage,
            vessel_name=vessel_name.strip(), voyage_no=voyage_no.strip(),
            invoice_container_count=1, container_count_basis="EXPLICIT",
        )
        sources = self.repository.list_sources(group_id)
        if sources:
            snapshots = list(
                self.inspect_sources(
                    probe,
                    bk_path=group.bk_path,
                    selections=[
                        ReconciliationSourceSelection(
                            item.bk_sheet,
                            (
                                vessel_name.strip()
                                if len(sources) == 1
                                else item.vessel_name
                            ),
                            (
                                voyage_no.strip()
                                if len(sources) == 1
                                else item.voyage_no
                            ),
                        )
                        for item in sources
                    ],
                )
            )
            snapshot = next(
                (item for item in snapshots if item.container_count), snapshots[0]
            )
            canonical_vessel_name = snapshot.canonical_vessel_name or vessel_name.strip()
            canonical_voyage_no = snapshot.canonical_voyage_no or voyage_no.strip()
            refreshed = self.repository.replace_group_sources(group_id, snapshots)
            issue = refreshed.bk_issue
        else:
            snapshot, issue = self._snapshot_for_period(
                probe, bk_path=group.bk_path, month=month, year=year
            )
            self._require_snapshot_match(probe, snapshot, issue=issue)
            canonical_vessel_name = snapshot.canonical_vessel_name or vessel_name.strip()
            canonical_voyage_no = snapshot.canonical_voyage_no or voyage_no.strip()
        for item in normalized:
            item["vessel_name"] = canonical_vessel_name
            item["voyage_no"] = canonical_voyage_no
            canonical_row = DataRow.from_mapping(item)
            item["fingerprint"] = self.contribution_fingerprint(
                canonical_row,
                source_sha256=str(item["source_sha256"]),
            )
        return self.repository.save_contributions(
            group_id,
            normalized,
            snapshot=None if sources else snapshot,
            vessel_name=canonical_vessel_name,
            voyage_no=canonical_voyage_no,
            month=month,
            year=year,
            bk_issue=issue,
        )

    def import_supplement(
        self,
        group_id: int,
        rows: Iterable[DataRow],
        *,
        source_batch_id: int | None,
        source_sha256: str,
    ) -> SupplementImportResult:
        group = self._require_group(group_id)
        if not group.is_current or group.status in {
            GroupStatus.ALLOCATED, GroupStatus.POSTED, GroupStatus.CANCELLED
        }:
            return SupplementImportResult(0, (), sum(1 for _ in rows), ("Hồ sơ đã hoàn tất.",))
        contributions = self.repository.list_contributions(group_id)
        existing_keys = {
            self._invoice_key(item.invoice_no)
            for item in contributions
            if self._invoice_key(item.invoice_no)
        }
        first = contributions[0] if contributions else None
        canonical_vessel_name = (
            str(getattr(first, "vessel_name", "") or "").strip()
            or group.vessel_key
        )
        canonical_voyage_no = (
            str(getattr(first, "voyage_no", "") or "").strip()
            or group.voyage_key
        )
        duplicates: list[str] = []
        errors: list[str] = []
        skipped = 0
        added = 0
        added_invoices: list[str] = []
        for index, row in enumerate(rows):
            if row.fee != "CB" or row.cont not in (None, ""):
                skipped += 1
                continue
            if not vessel_voyage_alias_equivalent(
                row.vessel_name,
                row.voyage_no,
                canonical_vessel_name,
                canonical_voyage_no,
            ):
                skipped += 1
                continue
            invoice_key = self._invoice_key(row.invoice_no)
            if not invoice_key:
                errors.append(f"Dòng {index + 1} chưa có số HĐ.")
                continue
            if invoice_key in existing_keys:
                duplicates.append(str(row.invoice_no).strip())
                continue
            try:
                self._require_candidate(row)
                self._add_row(
                    group_id,
                    row.copy_with(
                        vessel_name=canonical_vessel_name,
                        voyage_no=canonical_voyage_no,
                    ),
                    source_batch_id=source_batch_id,
                    source_item_index=index,
                    source_sha256=source_sha256,
                    source_kind="SUPPLEMENT",
                )
            except (DuplicateContributionError, SeaFreightMatchError, ValueError) as exc:
                errors.append(f"Dòng {index + 1}: {exc}")
                continue
            existing_keys.add(invoice_key)
            added += 1
            added_invoices.append(str(row.invoice_no).strip())
        return SupplementImportResult(
            added, tuple(duplicates), skipped, tuple(errors), tuple(added_invoices)
        )

    def prepare_allocation(self, group_id: int) -> list[DataRow]:
        group = self._require_group(group_id)
        contributions = self.repository.list_contributions(group_id)
        if not contributions:
            raise SeaFreightReconciliationError("Hồ sơ chưa có HĐ.")
        if group.status not in {GroupStatus.READY, GroupStatus.ALLOCATED}:
            raise SeaFreightReconciliationError("Số cont hóa đơn chưa khớp với BK.")
        first = contributions[0]
        sources = self.repository.list_sources(group_id)
        if sources:
            try:
                current_sources = self.inspect_sources(
                    self._contribution_row(first),
                    bk_path=group.bk_path,
                    selections=[
                        ReconciliationSourceSelection(
                            item.bk_sheet, item.vessel_name, item.voyage_no
                        )
                        for item in sources
                    ],
                )
            except Exception as exc:
                raise SeaFreightReconciliationError("Không đọc lại được dữ liệu BK.") from exc
            if any(
                expected.snapshot_hash != current.snapshot_hash
                for expected, current in zip(sources, current_sources, strict=True)
            ):
                self.repository.mark_needs_recheck(
                    group_id, "Một trong các sheet BK đã thay đổi trước khi xác nhận."
                )
                raise SeaFreightReconciliationError("BK đã thay đổi – cần kiểm tra lại.")
            if self.repository.unresolved_duplicate_containers(group_id):
                raise SeaFreightReconciliationError(
                    "Còn container trùng giữa các sheet chưa chọn nguồn."
                )
            return self.allocation_rows(group_id)
        try:
            current = self.matcher.snapshot(
                group.bk_path,
                group.bk_sheet,
                vessel_voyage_raw=first.vessel_voyage_raw,
                vessel_name=first.vessel_name,
                voyage_no=first.voyage_no,
            )
        except Exception as exc:
            raise SeaFreightReconciliationError("Không đọc lại được dữ liệu BK.") from exc
        if (
            group.bk_fingerprint != current.workbook_fingerprint
            or group.container_snapshot_hash != current.snapshot_hash
        ):
            self.repository.mark_needs_recheck(group_id, "BK thay đổi trước khi xác nhận.")
            raise SeaFreightReconciliationError("BK đã thay đổi – cần kiểm tra lại.")
        return self.allocation_rows(group_id)

    def create_revision(self, group_id: int) -> ReconciliationGroup:
        """Đọc lại BK và tạo một phiên mới từ hồ sơ đã hoàn tất."""

        group = self._require_group(group_id)
        if not group.is_current:
            raise SeaFreightReconciliationError(
                "Chỉ có thể đối soát lại từ phiên hiện hành."
            )
        if group.status not in {GroupStatus.ALLOCATED, GroupStatus.POSTED}:
            raise SeaFreightReconciliationError(
                "Hồ sơ chưa hoàn tất; hãy tiếp tục chỉnh sửa phiên hiện tại."
            )
        contributions = self.repository.list_contributions(group_id)
        if not contributions:
            raise SeaFreightReconciliationError("Hồ sơ không còn dữ liệu HĐ để chạy lại.")
        first = contributions[0]
        sources = self.repository.list_sources(group_id)
        if sources:
            try:
                snapshots = list(
                    self.inspect_sources(
                        self._contribution_row(first),
                        bk_path=group.bk_path,
                        selections=[
                            ReconciliationSourceSelection(
                                item.bk_sheet, item.vessel_name, item.voyage_no
                            )
                            for item in sources
                        ],
                    )
                )
            except Exception as exc:
                raise SeaFreightReconciliationError(
                    "Không đọc lại được dữ liệu BK để tạo phiên mới."
                ) from exc
            created = self.repository.create_revision(group_id, snapshots[0])
            return self.repository.replace_group_sources(created.id, snapshots)
        try:
            snapshot = self.matcher.snapshot(
                group.bk_path,
                group.bk_sheet,
                vessel_voyage_raw=first.vessel_voyage_raw,
                vessel_name=first.vessel_name,
                voyage_no=first.voyage_no,
            )
        except Exception as exc:
            raise SeaFreightReconciliationError(
                "Không đọc lại được dữ liệu BK để tạo phiên mới."
            ) from exc
        self._require_snapshot_match(self._contribution_row(first), snapshot)
        return self.repository.create_revision(group_id, snapshot)

    def allocation_rows(self, group_id: int) -> list[DataRow]:
        group = self._require_group(group_id)
        if group.status not in {
            GroupStatus.READY,
            GroupStatus.ALLOCATED,
            GroupStatus.POSTED,
        }:
            raise SeaFreightReconciliationError("Hồ sơ chưa đủ điều kiện xác nhận.")
        contributions = self.repository.list_contributions(group_id)
        containers = self.repository.list_container_rows(group_id)
        if not containers or len(containers) != group.bk_container_count:
            raise SeaFreightReconciliationError("Danh sách cont BK không còn đầy đủ.")
        invoice_no = ", ".join(self._unique(item.invoice_no for item in contributions)) or None
        bl = ", ".join(self._unique(item.bl for item in contributions)) or None
        carrier = " / ".join(self._unique(item.carrier for item in contributions)) or None
        date_values = [item.invoice_date for item in contributions]
        dates = self._unique(date_values)
        invoice_date = (
            dates[0]
            if len(dates) == 1 and all(value not in (None, "") for value in date_values)
            else None
        )
        amounts = allocate_integer_amount(group.total_amount, len(containers))
        first = contributions[0]
        source_document_id = f"SEA_RECON_{group.id}_R{group.revision_no}"
        source_document_name = (
            f"Cước biển {group.vessel_voyage_raw}"
            + (f" – HĐ {invoice_no}" if invoice_no else "")
        )
        return [
            DataRow(
                cont=str(container["container"]), bl=bl, fee="CB", rule="HD",
                amount=amounts[index], invoice_no=invoice_no, carrier=carrier,
                vessel_voyage_raw=group.vessel_voyage_raw,
                vessel_name=first.vessel_name, voyage_no=first.voyage_no,
                invoice_container_count=None, container_count_basis="UNKNOWN",
                invoice_date=invoice_date,
                source_document_id=source_document_id,
                source_document_name=source_document_name,
            )
            for index, container in enumerate(containers)
        ]

    @staticmethod
    def contribution_fingerprint(row: DataRow, *, source_sha256: str) -> str:
        payload = {
            "invoice_no": row.invoice_no,
            "invoice_date": row.invoice_date,
            "carrier": row.carrier,
            "bl": row.bl,
            "amount": row.amount,
            "vessel_name": normalize_match_key(row.vessel_name),
            "voyage_no": normalize_match_key(row.voyage_no),
            "invoice_container_count": row.invoice_container_count,
            "source_sha256": source_sha256,
        }
        return hashlib.sha256(
            json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
        ).hexdigest()

    @staticmethod
    def _require_candidate(row: DataRow) -> None:
        if row.fee != "CB" or row.cont not in (None, ""):
            raise SeaFreightMatchError("Chỉ nhận HĐ cước biển chưa có số cont.")
        vessel_voyage_keys(row.vessel_name, row.voyage_no)
        if type(row.invoice_container_count) is not int or row.invoice_container_count <= 0:
            raise SeaFreightMatchError("Số cont HĐ phải là số nguyên dương.")
        if row.container_count_basis not in {"EXPLICIT", "CALCULATED"}:
            raise SeaFreightMatchError("Căn cứ số cont chưa hợp lệ.")
        if type(row.amount) is not int or row.amount < 0:
            raise SeaFreightMatchError("Số tiền HĐ chưa hợp lệ.")

    def _require_snapshot_match(
        self,
        row: DataRow,
        snapshot: BkContainerSnapshot,
        *,
        issue: str | None = None,
    ) -> None:
        if snapshot.container_count > 0:
            return
        suggestions: tuple[VesselVoyageSuggestion, ...] = tuple(
            snapshot.alias_candidates
        )
        suggest = getattr(self.matcher, "suggestions", None)
        if not suggestions and callable(suggest) and issue != "SHEET_MISSING":
            try:
                suggestions = tuple(
                    suggest(
                        snapshot.bk_path,
                        snapshot.bk_sheet,
                        vessel_name=str(row.vessel_name or ""),
                        voyage_no=str(row.voyage_no or ""),
                        limit=5,
                    )
                )
            except Exception:
                suggestions = ()
        kind = snapshot.resolution_kind
        LOGGER.info(
            "Kết quả resolve tàu/chuyến: %s -> %s",
            kind.value if issue != "SHEET_MISSING" else "NOT_FOUND",
            vessel_voyage_text(row.vessel_name, row.voyage_no),
        )
        raise VesselVoyageNotFoundError(
            vessel_voyage=vessel_voyage_text(row.vessel_name, row.voyage_no),
            bk_sheet=snapshot.bk_sheet,
            suggestions=suggestions,
            sheet_missing=issue == "SHEET_MISSING",
            ambiguous=kind is VesselVoyageResolutionKind.AMBIGUOUS,
        )

    def _require_group(self, group_id: int) -> ReconciliationGroup:
        group = self.repository.get_group(group_id)
        if group is None:
            raise SeaFreightReconciliationError("Không tìm thấy hồ sơ đối soát.")
        return group

    @staticmethod
    def _contribution_row(item: object) -> DataRow:
        return DataRow(
            cont=None,
            bl=getattr(item, "bl"),
            fee="CB",
            rule="HD",
            amount=getattr(item, "amount"),
            invoice_no=getattr(item, "invoice_no"),
            carrier=getattr(item, "carrier"),
            vessel_voyage_raw=getattr(item, "vessel_voyage_raw"),
            vessel_name=getattr(item, "vessel_name"),
            voyage_no=getattr(item, "voyage_no"),
            invoice_container_count=getattr(item, "invoice_container_count"),
            container_count_basis=getattr(item, "container_count_basis"),
            invoice_date=getattr(item, "invoice_date"),
            source_document_id=getattr(item, "source_document_id", "LEGACY_DOCUMENT"),
            source_document_name=getattr(
                item, "source_document_name", "Dữ liệu bóc tách cũ"
            ),
        )

    @staticmethod
    def _invoice_key(value: object) -> str:
        return " ".join(str(value or "").split()).casefold()

    @staticmethod
    def _unique(values: Iterable[object]) -> list[str]:
        result: list[str] = []
        seen: set[str] = set()
        for value in values:
            text = " ".join(str(value).strip().split()) if value not in (None, "") else ""
            key = text.casefold()
            if not text or key in seen:
                continue
            seen.add(key)
            result.append(text)
        return result
