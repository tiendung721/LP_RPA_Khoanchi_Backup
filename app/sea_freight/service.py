from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Iterable

from app.models import DataRow

from .contracts import (
    BkContainerSnapshot,
    GroupStatus,
    ReconciliationGroup,
    SupplementImportResult,
    VesselVoyageSuggestion,
)
from .container_numbers import allocate_integer_amount
from .matching import (
    BkVesselMatcher,
    SeaFreightMatchError,
    normalize_match_key,
    vessel_voyage_keys,
    vessel_voyage_text,
)
from .repository import DuplicateContributionError, SeaFreightRepository


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
    ) -> None:
        self.vessel_voyage = vessel_voyage
        self.bk_sheet = bk_sheet
        self.suggestions = suggestions
        self.sheet_missing = sheet_missing
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
        snapshot, issue = self._snapshot_for_period(
            row, bk_path=bk_path, month=month, year=year
        )
        self._require_snapshot_match(row, snapshot, issue=issue)
        existing = self.repository.find_open_group(
            snapshot.bk_path, snapshot.bk_sheet, snapshot.vessel_key, snapshot.voyage_key
        )
        if existing is None:
            latest = self.repository.find_latest_group(
                snapshot.bk_path, snapshot.bk_sheet, snapshot.vessel_key, snapshot.voyage_key
            )
            if latest is not None and latest.status in {GroupStatus.ALLOCATED, GroupStatus.POSTED}:
                raise SeaFreightReconciliationError("Hồ sơ đã hoàn tất, không thể thêm HĐ.")
        if existing is not None:
            invoice_key = self._invoice_key(row.invoice_no)
            if invoice_key and invoice_key in {
                self._invoice_key(item.invoice_no)
                for item in self.repository.list_contributions(existing.id)
            }:
                return existing
            if existing.status in {GroupStatus.ALLOCATED, GroupStatus.POSTED}:
                raise SeaFreightReconciliationError("Hồ sơ đã hoàn tất, không thể thêm HĐ.")
        group = self.repository.upsert_snapshot(
            snapshot, month=month, year=year, bk_issue=issue
        )
        self._add_row(
            group.id,
            row,
            source_batch_id=source_batch_id,
            source_item_index=source_item_index,
            source_sha256=source_sha256,
            source_kind="INITIAL",
        )
        result = self.repository.get_group(group.id)
        assert result is not None
        return result

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
        self._add_row(
            group.id,
            row,
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
        month: int,
        year: int,
        source_batch_id: int | None,
        source_sha256: str,
    ) -> ReconciliationGroup:
        """Gom các HĐ cùng tàu/chuyến vào một hồ sơ bằng một lần ghi nguyên tử."""

        selected = list(rows)
        if not selected:
            raise SeaFreightReconciliationError("Chưa chọn hóa đơn cước biển.")
        first_index, first_row = selected[0]
        snapshot, issue = self._snapshot_for_period(
            first_row, bk_path=bk_path, month=month, year=year
        )
        self._require_snapshot_match(first_row, snapshot, issue=issue)
        for _, row in selected:
            self._require_candidate(row)
            if (
                normalize_match_key(row.vessel_name) != snapshot.vessel_key
                or normalize_match_key(row.voyage_no) != snapshot.voyage_key
            ):
                raise SeaFreightReconciliationError(
                    "Các hóa đơn được chọn không cùng tàu/chuyến."
                )
        group = self.repository.upsert_snapshot(
            snapshot, month=month, year=year, bk_issue=issue
        )
        existing_invoices = {
            self._invoice_key(item.invoice_no)
            for item in self.repository.list_contributions(group.id)
            if self._invoice_key(item.invoice_no)
        }
        payloads: list[dict[str, object]] = []
        for source_item_index, row in selected:
            invoice_key = self._invoice_key(row.invoice_no)
            if invoice_key and invoice_key in existing_invoices:
                continue
            payloads.append(
                self._row_payload(
                    row,
                    source_batch_id=source_batch_id,
                    source_item_index=source_item_index,
                    source_sha256=source_sha256,
                    source_kind="INITIAL",
                )
            )
            if invoice_key:
                existing_invoices.add(invoice_key)
        if payloads:
            self.repository.add_contributions(group.id, payloads)
        result = self.repository.get_group(group.id)
        assert result is not None
        return result

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
        snapshot, issue = self._snapshot_for_period(
            probe, bk_path=group.bk_path, month=month, year=year
        )
        self._require_snapshot_match(probe, snapshot, issue=issue)
        return self.repository.save_contributions(
            group_id,
            normalized,
            snapshot=snapshot,
            vessel_name=vessel_name.strip(),
            voyage_no=voyage_no.strip(),
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
        existing_keys = {
            self._invoice_key(item.invoice_no)
            for item in self.repository.list_contributions(group_id)
            if self._invoice_key(item.invoice_no)
        }
        duplicates: list[str] = []
        errors: list[str] = []
        skipped = 0
        added = 0
        added_invoices: list[str] = []
        for index, row in enumerate(rows):
            if row.fee != "CB" or row.cont not in (None, ""):
                skipped += 1
                continue
            if (
                normalize_match_key(row.vessel_name) != group.vessel_key
                or normalize_match_key(row.voyage_no) != group.voyage_key
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
                    row,
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
        suggestions: tuple[VesselVoyageSuggestion, ...] = ()
        suggest = getattr(self.matcher, "suggestions", None)
        if callable(suggest) and issue != "SHEET_MISSING":
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
        raise VesselVoyageNotFoundError(
            vessel_voyage=vessel_voyage_text(row.vessel_name, row.voyage_no),
            bk_sheet=snapshot.bk_sheet,
            suggestions=suggestions,
            sheet_missing=issue == "SHEET_MISSING",
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
