from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


class GroupStatus(str, Enum):
    PENDING = "PENDING"
    READY = "READY"
    OVER_COUNT = "OVER_COUNT"
    BK_NOT_FOUND = "BK_NOT_FOUND"
    METADATA_CONFLICT = "METADATA_CONFLICT"
    NEEDS_RECHECK = "NEEDS_RECHECK"
    ALLOCATED = "ALLOCATED"
    POSTED = "POSTED"
    CANCELLED = "CANCELLED"


class VesselVoyageResolutionKind(str, Enum):
    EXACT = "EXACT"
    AUTO_ALIAS = "AUTO_ALIAS"
    AMBIGUOUS = "AMBIGUOUS"
    NOT_FOUND = "NOT_FOUND"


@dataclass(frozen=True, slots=True)
class ContainerRecord:
    container: str
    source_sheet: str
    source_row: int
    source_sqt: int | None
    departure_date: str | None


@dataclass(frozen=True, slots=True)
class VesselVoyageSuggestion:
    vessel_voyage: str
    container_count: int
    score: float
    vessel_name: str = ""
    voyage_no: str = ""
    vessel_key: str = ""
    voyage_key: str = ""
    combined_key: str = ""
    selectable: bool = False


@dataclass(frozen=True, slots=True)
class VesselVoyageResolution:
    kind: VesselVoyageResolutionKind
    input_vessel_name: str
    input_voyage_no: str
    canonical_vessel_name: str = ""
    canonical_voyage_no: str = ""
    canonical_vessel_voyage: str = ""
    vessel_key: str = ""
    voyage_key: str = ""
    combined_key: str = ""
    candidates: tuple[VesselVoyageSuggestion, ...] = ()


@dataclass(frozen=True, slots=True)
class BkContainerSnapshot:
    bk_path: str
    bk_sheet: str
    vessel_voyage_raw: str
    vessel_key: str
    voyage_key: str
    combined_key: str
    workbook_fingerprint: str
    snapshot_hash: str
    containers: tuple[ContainerRecord, ...]
    invalid_container_count: int = 0
    duplicate_container_count: int = 0
    resolution_kind: VesselVoyageResolutionKind = VesselVoyageResolutionKind.EXACT
    canonical_vessel_name: str = ""
    canonical_voyage_no: str = ""
    alias_candidates: tuple[VesselVoyageSuggestion, ...] = ()

    @property
    def container_count(self) -> int:
        return len(self.containers)


@dataclass(frozen=True, slots=True)
class ReconciliationGroup:
    id: int
    bk_path: str
    bk_sheet: str
    vessel_voyage_raw: str
    vessel_key: str
    voyage_key: str
    combined_key: str
    bk_container_count: int
    received_container_count: int
    total_amount: int
    status: GroupStatus
    bk_fingerprint: str | None
    container_snapshot_hash: str | None
    output_carrier: str | None
    invoice_conflict_accepted: bool
    generated_batch_id: int | None
    posted_run_id: int | None
    created_at: str
    updated_at: str
    posted_at: str | None
    reconciliation_month: int | None = None
    reconciliation_year: int | None = None
    bk_issue: str | None = None
    bk_invalid_container_count: int = 0
    bk_duplicate_container_count: int = 0
    revision_no: int = 1
    supersedes_group_id: int | None = None
    is_current: bool = True
    primary_source_batch_id: int | None = None
    primary_source_item_index: int | None = None

    @property
    def missing_count(self) -> int:
        return max(0, self.bk_container_count - self.received_container_count)


@dataclass(frozen=True, slots=True)
class InvoiceContribution:
    id: int
    group_id: int
    source_batch_id: int | None
    source_item_index: int
    source_sha256: str
    invoice_no: str | None
    invoice_date: str | None
    bl: str | None
    vessel_voyage_raw: str
    vessel_name: str
    voyage_no: str
    invoice_container_count: int
    container_count_basis: str
    carrier: str | None
    amount: int
    fingerprint: str
    status: str
    created_at: str
    updated_at: str
    removed_at: str | None
    source_kind: str = "INITIAL"
    source_document_id: str = "LEGACY_DOCUMENT"
    source_document_name: str = "Dữ liệu bóc tách cũ"


class InvoiceHistoryMatchKind(str, Enum):
    NONE = "NONE"
    EXACT = "EXACT"
    CONFLICT = "CONFLICT"


@dataclass(frozen=True, slots=True)
class InvoiceHistoryMatch:
    kind: InvoiceHistoryMatchKind
    group: ReconciliationGroup | None = None
    contribution: InvoiceContribution | None = None
    differing_fields: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class ReconciliationOpenResult:
    group: ReconciliationGroup
    added_source_indices: tuple[int, ...] = ()
    duplicate_source_indices: tuple[int, ...] = ()
    requires_revision: bool = False
    resolution_kind: VesselVoyageResolutionKind = VesselVoyageResolutionKind.EXACT
    canonical_vessel_name: str = ""
    canonical_voyage_no: str = ""


@dataclass(frozen=True, slots=True)
class SupplementImportResult:
    added_count: int
    duplicate_invoices: tuple[str, ...]
    skipped_count: int
    errors: tuple[str, ...]
    added_invoices: tuple[str, ...] = ()


GROUP_STATUS_LABELS = {
    GroupStatus.PENDING: "Thiếu dữ liệu hóa đơn",
    GroupStatus.READY: "Đủ số cont – chờ xác nhận",
    GroupStatus.OVER_COUNT: "Số cont hóa đơn đang thừa",
    GroupStatus.BK_NOT_FOUND: "Chưa tìm thấy dữ liệu BK",
    GroupStatus.METADATA_CONFLICT: "Cần sửa dữ liệu hóa đơn",
    GroupStatus.NEEDS_RECHECK: "BK đã thay đổi – cần kiểm tra lại",
    GroupStatus.ALLOCATED: "Đã hoàn tất đối soát",
    GroupStatus.POSTED: "Đã ghi vào BK",
    GroupStatus.CANCELLED: "Đã ngừng sử dụng",
}


def group_status_text(group: ReconciliationGroup) -> str:
    """Nhãn tiếng Việt duy nhất dùng cho mọi màn hình đối soát."""

    if group.status is GroupStatus.BK_NOT_FOUND:
        if group.bk_issue == "SHEET_MISSING" and group.reconciliation_month and group.reconciliation_year:
            return f"Chưa có BK tháng {group.reconciliation_month:02d}/{group.reconciliation_year}"
        return "Chưa tìm thấy tàu/chuyến"
    if group.status is GroupStatus.PENDING:
        return f"Thiếu {group.missing_count} cont"
    if group.status is GroupStatus.OVER_COUNT:
        return f"Thừa {group.received_container_count - group.bk_container_count} cont"
    if group.status is GroupStatus.READY:
        return f"Đủ {group.received_container_count}/{group.bk_container_count} cont – chờ xác nhận"
    return GROUP_STATUS_LABELS.get(group.status, "Cần kiểm tra")
