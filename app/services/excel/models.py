"""Pure-Python contracts used by the Excel workflows.

The UI deliberately consumes these value objects instead of importing openpyxl.
Keeping the plans serialisable-ish also makes conflict dialogs and audit logging
straightforward.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Mapping, Sequence


class ExcelOperation(str, Enum):
    DAILY_SYNC = "DAILY_SYNC"
    EXPENSE_POSTING = "EXPENSE_POSTING"
    PAYMENT_SYNC = "PAYMENT_SYNC"


class ExcelRunStatus(str, Enum):
    ANALYZING = "ANALYZING"
    WAITING_USER = "WAITING_USER"
    APPLYING = "APPLYING"
    SUCCEEDED = "SUCCEEDED"
    NO_CHANGES = "NO_CHANGES"
    CANCELLED = "CANCELLED"
    FAILED = "FAILED"


class OutcomeStatus(str, Enum):
    """Kết quả thực thi thực tế của một item hoặc một thành phần dữ liệu."""

    WRITTEN = "WRITTEN"
    PARTIAL = "PARTIAL"
    UNCHANGED = "UNCHANGED"
    USER_KEPT = "USER_KEPT"
    USER_SKIPPED = "USER_SKIPPED"
    INVALID_SOURCE = "INVALID_SOURCE"
    FAILED = "FAILED"


@dataclass(slots=True)
class FieldWriteOutcome:
    field_name: str
    source_value: Any = None
    target_value_before: Any = None
    target_value_after: Any = None
    action: str | None = None
    status: OutcomeStatus = OutcomeStatus.UNCHANGED
    source_sheet: str | None = None
    source_row: int | None = None
    target_sheet: str | None = None
    target_row: int | None = None
    target_cell: str | None = None
    reason_code: str | None = None
    reason: str = ""


@dataclass(slots=True)
class ItemWriteOutcome:
    item_id: str
    source_index: int | None = None
    container: str | None = None
    sqt: int | None = None
    bl: str | None = None
    fee: str | None = None
    source_sheet: str | None = None
    source_row: int | None = None
    target_sheet: str | None = None
    target_row: int | None = None
    status: OutcomeStatus = OutcomeStatus.UNCHANGED
    fields: list[FieldWriteOutcome] = field(default_factory=list)


class PostingItemStatus(str, Enum):
    PLANNED = "PLANNED"
    POSTED = "POSTED"
    ALREADY_EXISTS = "ALREADY_EXISTS"
    USER_SKIPPED = "USER_SKIPPED"
    NOT_MATCHED = "NOT_MATCHED"
    UNRESOLVED = "UNRESOLVED"
    FAILED = "FAILED"


class ConflictType(str, Enum):
    INVALID_SQT = "INVALID_SQT"
    DUPLICATE_SOURCE_ROW = "DUPLICATE_SOURCE_ROW"
    SYNC_GROUP_COUNT_MISMATCH = "SYNC_GROUP_COUNT_MISMATCH"
    TARGET_MONTH_AMBIGUOUS = "TARGET_MONTH_AMBIGUOUS"
    TARGET_SHEET_AMBIGUOUS = "TARGET_SHEET_AMBIGUOUS"
    CONTAINER_NOT_FOUND = "CONTAINER_NOT_FOUND"
    MULTIPLE_CONTAINER_MATCH = "MULTIPLE_CONTAINER_MATCH"
    REPEATED_SOURCE_CONTAINER = "REPEATED_SOURCE_CONTAINER"
    TARGET_CELL_OCCUPIED = "TARGET_CELL_OCCUPIED"
    TARGET_CELL_FORMULA = "TARGET_CELL_FORMULA"
    TARGET_CELL_TEXT = "TARGET_CELL_TEXT"
    UNKNOWN_FEE_CODE = "UNKNOWN_FEE_CODE"
    FEE_COLUMN_MISSING = "FEE_COLUMN_MISSING"
    BL_ONLY_NO_CONTAINER = "BL_ONLY_NO_CONTAINER"
    BATCH_ALREADY_POSTED = "BATCH_ALREADY_POSTED"
    FILE_CHANGED = "FILE_CHANGED"
    FILE_LOCKED = "FILE_LOCKED"
    PARTIAL_KEY_MATCH = "PARTIAL_KEY_MATCH"
    NEGATIVE_ADJUSTMENT = "NEGATIVE_ADJUSTMENT"
    PAYMENT_SOURCE_INVALID = "PAYMENT_SOURCE_INVALID"
    PAYMENT_CLEAR_VALUE = "PAYMENT_CLEAR_VALUE"
    MULTIPLE_SOURCE_INVOICES = "MULTIPLE_SOURCE_INVOICES"
    INVOICE_VALUE_CONFLICT = "INVOICE_VALUE_CONFLICT"
    INVOICE_COLUMN_MISSING = "INVOICE_COLUMN_MISSING"
    MULTIPLE_EXPENSE_SAME_CELL = "MULTIPLE_EXPENSE_SAME_CELL"
    CARRY_FORWARD_MAPPING_INVALID = "CARRY_FORWARD_MAPPING_INVALID"
    MULTIPLE_SOURCE_CARRIERS = "MULTIPLE_SOURCE_CARRIERS"
    CARRIER_VALUE_CONFLICT = "CARRIER_VALUE_CONFLICT"
    CARRIER_COLUMN_MISSING = "CARRIER_COLUMN_MISSING"


class ResolutionAction(str, Enum):
    SKIP_INVALID = "SKIP_INVALID"
    KEEP_ONE = "KEEP_ONE"
    KEEP_ALL = "KEEP_ALL"
    CANCEL_ALL = "CANCEL_ALL"
    SELECT_MONTH = "SELECT_MONTH"
    SELECT_SHEET = "SELECT_SHEET"
    SELECT_ROW = "SELECT_ROW"
    SELECT_FEE = "SELECT_FEE"
    SELECT_INVOICE = "SELECT_INVOICE"
    SELECT_SOURCE_ITEM = "SELECT_SOURCE_ITEM"
    SELECT_CARRIER = "SELECT_CARRIER"
    KEEP_EXISTING = "KEEP_EXISTING"
    KEEP_FORMULA = "KEEP_FORMULA"
    OVERWRITE = "OVERWRITE"
    ADD = "ADD"
    APPEND_CARRIER = "APPEND_CARRIER"
    SKIP = "SKIP"
    SKIP_INVOICE = "SKIP_INVOICE"
    POST_UNPOSTED_ONLY = "POST_UNPOSTED_ONLY"
    REANALYZE = "REANALYZE"
    RETRY = "RETRY"
    CANCEL = "CANCEL"


class TargetCellKind(str, Enum):
    EMPTY = "EMPTY"
    ZERO = "ZERO"
    SAME_VALUE = "SAME_VALUE"
    NUMBER = "NUMBER"
    FORMULA = "FORMULA"
    TEXT = "TEXT"


class SyncActionType(str, Enum):
    INSERT = "INSERT"
    UPDATE = "UPDATE"
    UNCHANGED = "UNCHANGED"
    TARGET_ONLY = "TARGET_ONLY"


@dataclass(frozen=True, slots=True)
class WorkbookFingerprint:
    size: int
    mtime_ns: int
    sha256: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "size": self.size,
            "mtime_ns": self.mtime_ns,
            "sha256": self.sha256,
        }

    @classmethod
    def from_value(
        cls, value: "WorkbookFingerprint | Mapping[str, Any]"
    ) -> "WorkbookFingerprint":
        if isinstance(value, cls):
            return value
        return cls(
            size=int(value["size"]),
            mtime_ns=int(value["mtime_ns"]),
            sha256=str(value["sha256"]),
        )


@dataclass(frozen=True, slots=True)
class TargetCellState:
    kind: TargetCellKind
    value: Any
    coordinate: str
    data_type: str | None = None


@dataclass(slots=True)
class SourceSheetCandidate:
    month: int
    source_sheet: str

    @property
    def sheet_name(self) -> str:
        return self.source_sheet


@dataclass(slots=True)
class MonthCandidate:
    month: int
    source_sheet: str
    target_sheet: str
    new_row_count: int = 0
    update_count: int = 0
    unchanged_count: int = 0
    target_only_count: int = 0
    invalid_count: int = 0
    last_sqt: int = 0
    match_count: int = 0
    recently_synced: bool = False
    year: int | None = None

    @property
    def sheet_name(self) -> str:
        return self.target_sheet


@dataclass(slots=True)
class SyncRow:
    source_sheet: str
    source_row: int
    month: int
    sqt: int
    values: tuple[Any, ...]
    duplicate_key: tuple[Any, ...]

    @property
    def container(self) -> Any:
        return self.values[2] if len(self.values) > 2 else None


@dataclass(slots=True)
class SyncAction:
    action: SyncActionType
    sqt: int
    source: SyncRow | None = None
    target_row: int | None = None
    target_values: tuple[Any, ...] | None = None
    protected_values: dict[int, Any] = field(default_factory=dict)


@dataclass(slots=True)
class SyncConflict:
    conflict_id: str
    conflict_type: ConflictType
    message: str
    source_sheet: str | None = None
    source_row: int | None = None
    sqt: int | None = None
    container: str | None = None
    allowed_actions: tuple[ResolutionAction, ...] = ()
    default_action: ResolutionAction | None = None
    details: dict[str, Any] = field(default_factory=dict)

    @property
    def id(self) -> str:
        return self.conflict_id

    @property
    def type(self) -> ConflictType:
        return self.conflict_type

    @property
    def options(self) -> tuple[ResolutionAction, ...]:
        return self.allowed_actions

    @property
    def default_resolution(self) -> ResolutionAction | None:
        return self.default_action


@dataclass(slots=True)
class SyncResolution:
    conflict_id: str
    action: ResolutionAction | str
    selected_month: int | None = None
    selected_sheet: str | None = None


@dataclass(slots=True)
class SyncPlan:
    source_path: Path
    target_path: Path
    source_fingerprint: WorkbookFingerprint
    target_fingerprint: WorkbookFingerprint
    month_candidates: list[MonthCandidate]
    rows_by_month: dict[int, list[SyncRow]]
    source_year: int | None = None
    target_year: int | None = None
    rows_by_target: dict[str, list[SyncRow]] = field(default_factory=dict)
    actions_by_target: dict[str, list[SyncAction]] = field(default_factory=dict)
    invalid_rows: list[int] = field(default_factory=list)
    source_snapshot_path: Path | None = None
    source_snapshot_fingerprint: WorkbookFingerprint | None = None
    conflicts: list[SyncConflict] = field(default_factory=list)
    selected_month: int | None = None
    selected_target_sheet: str | None = None
    run_id: int | None = None
    source_target_sheets: dict[str, str | None] = field(default_factory=dict)
    selected_target_sheets: set[str] = field(default_factory=set)

    @property
    def operation(self) -> ExcelOperation:
        return ExcelOperation.DAILY_SYNC

    @property
    def rows(self) -> list[SyncRow]:
        if self.selected_target_sheet is not None and self.rows_by_target:
            return list(self.rows_by_target.get(self.selected_target_sheet, ()))
        if self.selected_month is None:
            return []
        return list(self.rows_by_month.get(self.selected_month, ()))

    @property
    def selected_sheet(self) -> str | None:
        if self.selected_target_sheet is not None:
            return self.selected_target_sheet
        for candidate in self.month_candidates:
            if candidate.month == self.selected_month:
                return candidate.target_sheet
        return None

    @property
    def has_changes(self) -> bool:
        return any(
            action.action in {SyncActionType.INSERT, SyncActionType.UPDATE}
            for actions in self.actions_by_target.values()
            for action in actions
        )

    @property
    def requires_user_input(self) -> bool:
        return (
            (
                any(target is None for target in self.source_target_sheets.values())
                if self.source_target_sheets
                else self.selected_sheet is None
            )
            or self.has_changes
            or any(conflict.default_action is None for conflict in self.conflicts)
        )

    @property
    def actions(self) -> list[SyncAction]:
        if self.selected_target_sheets:
            return [
                action
                for sheet in self.selected_target_sheets
                for action in self.actions_by_target.get(sheet, ())
            ]
        sheet = self.selected_sheet
        return list(self.actions_by_target.get(sheet, ())) if sheet else []

    def _count(self, action: SyncActionType) -> int:
        return sum(item.action is action for item in self.actions)

    @property
    def insert_count(self) -> int:
        return self._count(SyncActionType.INSERT)

    @property
    def update_count(self) -> int:
        return self._count(SyncActionType.UPDATE)

    @property
    def unchanged_count(self) -> int:
        return self._count(SyncActionType.UNCHANGED)

    @property
    def target_only_count(self) -> int:
        return self._count(SyncActionType.TARGET_ONLY)

    @property
    def invalid_count(self) -> int:
        return len(self.invalid_rows)

    @property
    def conflict_count(self) -> int:
        return sum(
            conflict.default_action is None for conflict in self.conflicts
        )


@dataclass(slots=True)
class SyncResult:
    status: ExcelRunStatus
    target_path: Path
    sheet_name: str | None = None
    added_rows: int = 0
    inserted_rows: int = 0
    updated_rows: int = 0
    unchanged_rows: int = 0
    target_only_rows: int = 0
    invalid_rows: int = 0
    skipped_rows: int = 0
    conflict_count: int = 0
    backup_path: Path | None = None
    fingerprint_before: WorkbookFingerprint | None = None
    fingerprint_after: WorkbookFingerprint | None = None
    run_id: int | None = None
    message: str = ""
    target_sheets: tuple[str, ...] = ()
    item_outcomes: list[ItemWriteOutcome] = field(default_factory=list)

    @property
    def operation(self) -> ExcelOperation:
        return ExcelOperation.DAILY_SYNC


@dataclass(slots=True)
class RowCandidate:
    row: int
    sqt: int | None
    container: str | None
    source_sheet: str = ""
    source_row: int | None = None
    plan_values: tuple[Any, ...] = ()
    source_signature: str = ""
    carried_target_row: int | None = None
    mapping_invalid: bool = False
    cargo_type: Any = None
    closing_date: Any = None
    vessel: Any = None
    recipient: Any = None
    carrier: Any = None
    is_ron: bool = False


@dataclass(slots=True)
class PostingItem:
    source_indices: list[int]
    container: str | None
    bl: str | None
    original_fee: str
    selected_fee: str
    rule: str | None
    amount: int
    vessel_voyage_raw: str | None = None
    vessel_name: str | None = None
    voyage_no: str | None = None
    sheet_name: str | None = None
    selected_source_sheet: str | None = None
    selected_source_row: int | None = None
    source_sqt: int | None = None
    plan_values: tuple[Any, ...] = ()
    source_signature: str = ""
    carry_forward_required: bool = False
    target_row: int | None = None
    target_column: int | None = None
    target_cell: str | None = None
    current_value: Any = None
    cell_state: TargetCellState | None = None
    action: ResolutionAction | None = None
    status: PostingItemStatus = PostingItemStatus.PLANNED
    row_candidates: list[RowCandidate] = field(default_factory=list)
    source_items: list[dict[str, Any]] = field(default_factory=list)
    force_repost: bool = False
    invoice_candidates: list[str] = field(default_factory=list)
    selected_invoice: str | None = None
    invoice_column: int | None = None
    invoice_cell: str | None = None
    invoice_current_value: Any = None
    invoice_value_after: Any = None
    invoice_action: ResolutionAction | None = None
    carrier_candidates: list[str] = field(default_factory=list)
    selected_carrier: str | None = None
    carrier_group: str | None = None
    carrier_column: int | None = None
    carrier_cell: str | None = None
    carrier_current_value: Any = None
    carrier_value_after: Any = None
    carrier_action: ResolutionAction | None = None

    @property
    def fee(self) -> str:
        return self.selected_fee

    @property
    def row(self) -> int | None:
        return self.target_row

    @property
    def column(self) -> int | None:
        return self.target_column

    @property
    def cell(self) -> str | None:
        return self.target_cell


@dataclass(slots=True)
class PostingSourceGroup:
    group_id: str
    source_document_id: str
    source_document_name: str
    invoice_no: str | None
    source_item_indices: list[int]
    invoice_dates: list[str]
    suggested_sheet: str | None = None
    target_sheet: str | None = None
    reconciliation_group_id: int | None = None
    target_locked: bool = False
    can_split_by_invoice: bool = False

@dataclass(slots=True)
class PostingConflict:
    conflict_id: str
    conflict_type: ConflictType
    message: str
    item_index: int | None = None
    container: str | None = None
    bl: str | None = None
    sqt: int | None = None
    fee: str | None = None
    amount: int | None = None
    carrier: str | None = None
    vessel_voyage: str | None = None
    sheet_name: str | None = None
    target_row: int | None = None
    target_column: int | None = None
    target_cell: str | None = None
    current_value: Any = None
    allowed_actions: tuple[ResolutionAction, ...] = ()
    default_action: ResolutionAction | None = None
    row_candidates: list[RowCandidate] = field(default_factory=list)
    details: dict[str, Any] = field(default_factory=dict)

    @property
    def id(self) -> str:
        return self.conflict_id

    @property
    def type(self) -> ConflictType:
        return self.conflict_type

    @property
    def options(self) -> tuple[ResolutionAction, ...]:
        return self.allowed_actions

    @property
    def default_resolution(self) -> ResolutionAction | None:
        return self.default_action


@dataclass(slots=True)
class PostingResolution:
    conflict_id: str
    action: ResolutionAction | str
    selected_sheet: str | None = None
    selected_row: int | None = None
    selected_source_sheet: str | None = None
    selected_source_item_index: int | None = None
    selected_fee: str | None = None
    selected_invoice: str | None = None
    selected_carrier: str | None = None


@dataclass(slots=True)
class PostingPlan:
    batch_id: int | None
    batch_path: Path
    batch_hash: str
    target_path: Path
    target_fingerprint: WorkbookFingerprint
    items: list[PostingItem]
    conflicts: list[PostingConflict]
    sheet_candidates: list[MonthCandidate]
    selected_sheet: str | None = None
    source_item_count: int = 0
    already_posted_indices: set[int] = field(default_factory=set)
    previously_posted_items: list[Any] = field(default_factory=list)
    repost_source_indices: set[int] = field(default_factory=set)
    repost_selection_done: bool = False
    run_id: int | None = None
    source_members: list[dict[str, Any]] = field(default_factory=list)
    reconciliation_members: list[dict[str, Any]] = field(default_factory=list)
    original_source_count: int = 0
    reconciliation_source_count: int = 0
    confirmation_required: bool = False
    confirmation_done: bool = False
    source_kind: str = "ASSISTANT"
    target_sheets: set[str] = field(default_factory=set)
    source_groups: list[PostingSourceGroup] = field(default_factory=list)
    split_document_ids: set[str] = field(default_factory=set)

    @property
    def operation(self) -> ExcelOperation:
        return ExcelOperation.EXPENSE_POSTING

    @property
    def requires_user_input(self) -> bool:
        return (
            (self.confirmation_required and not self.confirmation_done)
            or
            (
                self.source_kind != "BANG_KE"
                and (
                    any(group.target_sheet is None for group in self.source_groups)
                    if self.source_groups
                    else self.selected_sheet is None
                )
            )
            or (
                bool(self.previously_posted_items)
                and not self.repost_selection_done
            )
            or bool(self.conflicts)
        )

    @property
    def has_changes(self) -> bool:
        return any(
            item.status is PostingItemStatus.PLANNED
            and item.action not in {
                ResolutionAction.SKIP,
                ResolutionAction.KEEP_EXISTING,
                ResolutionAction.KEEP_FORMULA,
            }
            for item in self.items
        )


@dataclass(slots=True)
class PostingResult:
    status: ExcelRunStatus
    target_path: Path
    sheet_name: str | None = None
    target_sheets: tuple[str, ...] = ()
    posted_source_items: int = 0
    written_cells: int = 0
    invoice_written_cells: int = 0
    carrier_written_cells: int = 0
    carrier_appended_cells: int = 0
    carrier_kept_cells: int = 0
    unmapped_carrier_items: int = 0
    skipped_source_items: int = 0
    already_existing_items: int = 0
    conflict_count: int = 0
    backup_path: Path | None = None
    fingerprint_before: WorkbookFingerprint | None = None
    fingerprint_after: WorkbookFingerprint | None = None
    run_id: int | None = None
    message: str = ""
    item_outcomes: list[ItemWriteOutcome] = field(default_factory=list)

    @property
    def operation(self) -> ExcelOperation:
        return ExcelOperation.EXPENSE_POSTING


@dataclass(slots=True)
class PaymentSyncItem:
    item_id: str
    source_row: int
    source_rows: tuple[int, ...]
    sqt: int
    container: str
    values: dict[str, Any]
    target_type: str = "NAM"
    target_row: int | None = None
    # Dòng do user chọn phải độc lập với dòng được thuật toán tự khớp. Giá trị
    # này được giữ xuyên suốt các vòng refine và luôn được ưu tiên khi phân tích.
    selected_target_row: int | None = None
    write_identity: bool = False
    skip_item: bool = False
    status: str = "NEW"
    differences: dict[str, tuple[Any, Any]] = field(default_factory=dict)
    amount_actions: dict[str, ResolutionAction] = field(default_factory=dict)
    amount_review_values: dict[str, tuple[Any, Any]] = field(default_factory=dict)
    invoice_values: dict[str, str | None] = field(default_factory=dict)
    invoice_candidates: dict[str, list[str]] = field(default_factory=dict)
    selected_invoices: dict[str, str] = field(default_factory=dict)
    invoice_actions: dict[str, ResolutionAction] = field(default_factory=dict)
    invoice_differences: dict[str, tuple[Any, str]] = field(default_factory=dict)
    invoice_review_values: dict[str, tuple[Any, str]] = field(default_factory=dict)
    carrier_value: str | None = None
    carrier_action: ResolutionAction | None = None
    carrier_difference: tuple[Any, str] | None = None
    carrier_review_value: tuple[Any, str] | None = None

    @property
    def is_new(self) -> bool:
        return self.status == "NEW"

    @property
    def is_update(self) -> bool:
        return self.status == "UPDATE"

    @property
    def is_unchanged(self) -> bool:
        return self.status == "UNCHANGED"


@dataclass(slots=True)
class PaymentSyncConflict:
    conflict_id: str
    conflict_type: ConflictType
    message: str
    item_id: str
    source_row: int
    sqt: int
    container: str
    carrier: str | None = None
    allowed_actions: tuple[ResolutionAction, ...] = (
        ResolutionAction.SKIP,
        ResolutionAction.SELECT_ROW,
        ResolutionAction.CANCEL_ALL,
    )
    default_action: ResolutionAction = ResolutionAction.SKIP
    row_candidates: list[RowCandidate] = field(default_factory=list)
    fee: str | None = None
    sheet_name: str | None = None
    target_row: int | None = None
    target_column: int | None = None
    target_cell: str | None = None
    current_value: Any = None
    details: dict[str, Any] = field(default_factory=dict)

    @property
    def id(self) -> str:
        return self.conflict_id

    @property
    def type(self) -> ConflictType:
        return self.conflict_type


@dataclass(slots=True)
class PaymentTargetPlan:
    target_type: str
    sheet_name: str
    items: list[PaymentSyncItem]
    conflicts: list[PaymentSyncConflict] = field(default_factory=list)
    sheet_to_create: bool = False
    template_sheet: str | None = None

    @property
    def new_rows(self) -> list[PaymentSyncItem]:
        return [item for item in self.items if item.is_new]

    @property
    def update_rows(self) -> list[PaymentSyncItem]:
        return [item for item in self.items if item.is_update]

    @property
    def unchanged_rows(self) -> list[PaymentSyncItem]:
        return [item for item in self.items if item.is_unchanged]

    @property
    def new_count(self) -> int:
        return len(self.new_rows)

    @property
    def update_count(self) -> int:
        return len(self.update_rows)

    @property
    def unchanged_count(self) -> int:
        return len(self.unchanged_rows)

    @property
    def conflict_count(self) -> int:
        return len(self.conflicts)

    @property
    def invoice_change_count(self) -> int:
        return sum(len(item.invoice_differences) for item in self.items)


@dataclass(slots=True)
class PaymentSyncPlan:
    source_path: Path
    target_path: Path
    source_fingerprint: WorkbookFingerprint
    target_fingerprint: WorkbookFingerprint
    source_sheet: str
    targets: dict[str, PaymentTargetPlan]
    normalization_required: bool = False
    normalization_sheet_count: int = 0
    source_vba_present: bool = False
    target_vba_present: bool = False
    run_id: int | None = None

    @property
    def operation(self) -> ExcelOperation:
        return ExcelOperation.PAYMENT_SYNC

    @property
    def selected_sheet(self) -> str:
        return ", ".join(target.sheet_name for target in self.targets.values())

    @property
    def target_sheet(self) -> str:
        return self.selected_sheet

    @property
    def items(self) -> list[PaymentSyncItem]:
        return [item for target in self.targets.values() for item in target.items]

    @property
    def conflicts(self) -> list[PaymentSyncConflict]:
        return [
            conflict
            for target in self.targets.values()
            for conflict in target.conflicts
        ]

    @property
    def target_sheet_created(self) -> bool:
        return any(target.sheet_to_create for target in self.targets.values())

    @property
    def template_sheet(self) -> str | None:
        templates = [
            target.template_sheet
            for target in self.targets.values()
            if target.template_sheet
        ]
        return ", ".join(templates) if templates else None

    @property
    def new_rows(self) -> list[PaymentSyncItem]:
        return [item for item in self.items if item.is_new]

    @property
    def update_rows(self) -> list[PaymentSyncItem]:
        return [item for item in self.items if item.is_update]

    @property
    def unchanged_rows(self) -> list[PaymentSyncItem]:
        return [item for item in self.items if item.is_unchanged]

    @property
    def new_count(self) -> int:
        return len(self.new_rows)

    @property
    def update_count(self) -> int:
        return len(self.update_rows)

    @property
    def unchanged_count(self) -> int:
        return len(self.unchanged_rows)

    @property
    def conflict_count(self) -> int:
        return len(self.conflicts)

    @property
    def invoice_change_count(self) -> int:
        return sum(target.invoice_change_count for target in self.targets.values())

    @property
    def has_changes(self) -> bool:
        return (
            self.target_sheet_created
            or self.normalization_required
            or bool(self.new_rows)
            or bool(self.update_rows)
        )

    @property
    def requires_user_input(self) -> bool:
        return bool(
            self.target_sheet_created
            or self.conflicts
            or self.new_rows
            or self.update_rows
        )


@dataclass(slots=True)
class PaymentSyncBatchPlan:
    source_path: Path
    target_path: Path
    source_fingerprint: WorkbookFingerprint
    target_fingerprint: WorkbookFingerprint
    month_plans: dict[str, PaymentSyncPlan]
    normalization_required: bool = False
    normalization_sheet_count: int = 0
    source_vba_present: bool = False
    target_vba_present: bool = False
    run_id: int | None = None

    @property
    def operation(self) -> ExcelOperation:
        return ExcelOperation.PAYMENT_SYNC

    @property
    def source_sheet(self) -> str:
        return ", ".join(self.month_plans)

    @property
    def selected_sheet(self) -> str:
        return ", ".join(
            target.sheet_name
            for plan in self.month_plans.values()
            for target in plan.targets.values()
        )

    target_sheet = selected_sheet

    @property
    def items(self) -> list[PaymentSyncItem]:
        return [item for plan in self.month_plans.values() for item in plan.items]

    @property
    def conflicts(self) -> list[PaymentSyncConflict]:
        return [
            conflict
            for plan in self.month_plans.values()
            for conflict in plan.conflicts
        ]

    @property
    def new_rows(self) -> list[PaymentSyncItem]:
        return [item for plan in self.month_plans.values() for item in plan.new_rows]

    @property
    def update_rows(self) -> list[PaymentSyncItem]:
        return [item for plan in self.month_plans.values() for item in plan.update_rows]

    @property
    def unchanged_rows(self) -> list[PaymentSyncItem]:
        return [item for plan in self.month_plans.values() for item in plan.unchanged_rows]

    @property
    def new_count(self) -> int:
        return len(self.new_rows)

    @property
    def update_count(self) -> int:
        return len(self.update_rows)

    @property
    def unchanged_count(self) -> int:
        return len(self.unchanged_rows)

    @property
    def conflict_count(self) -> int:
        return len(self.conflicts)

    @property
    def target_sheet_created(self) -> bool:
        return any(plan.target_sheet_created for plan in self.month_plans.values())

    @property
    def invoice_change_count(self) -> int:
        return sum(plan.invoice_change_count for plan in self.month_plans.values())

    @property
    def template_sheet(self) -> str | None:
        names = [
            target.template_sheet
            for plan in self.month_plans.values()
            for target in plan.targets.values()
            if target.template_sheet
        ]
        return ", ".join(names) if names else None

    @property
    def has_changes(self) -> bool:
        return self.normalization_required or any(
            plan.has_changes for plan in self.month_plans.values()
        )

    @property
    def requires_user_input(self) -> bool:
        return any(plan.requires_user_input for plan in self.month_plans.values())


@dataclass(slots=True)
class PaymentTargetResult:
    target_type: str
    sheet_name: str
    sheet_created: bool = False
    template_sheet_name: str | None = None
    inserted_rows: int = 0
    updated_rows: int = 0
    unchanged_rows: int = 0
    skipped_rows: int = 0
    conflict_count: int = 0
    invoice_written_cells: int = 0
    carrier_written_cells: int = 0
    carrier_appended_cells: int = 0
    carrier_kept_cells: int = 0
    unmapped_carrier_items: int = 0


@dataclass(slots=True)
class CarrierSummaryResult:
    """Kết quả dựng bảng tổng hợp vận tải từ sổ chi tiết tích lũy."""

    changed: bool = False
    source_sheet_name: str = ""
    period: str = ""
    carrier_count: int = 0
    invoice_count: int = 0
    selected_period_total: int | float = 0
    missing_invoice_items: int = 0
    unmapped_carrier_items: int = 0
    cross_carrier_invoices: int = 0
    suspected_duplicate_items: int = 0
    detail_count: int = 0


@dataclass(slots=True)
class PaymentSyncResult:
    status: ExcelRunStatus
    source_path: Path
    target_path: Path
    target_results: dict[str, PaymentTargetResult]
    source_sheet_name: str
    backup_path: Path | None = None
    source_backup_path: Path | None = None
    fingerprint_before: WorkbookFingerprint | None = None
    fingerprint_after: WorkbookFingerprint | None = None
    source_fingerprint_after: WorkbookFingerprint | None = None
    vba_preserved: bool = True
    carrier_summary: CarrierSummaryResult | None = None
    run_id: int | None = None
    message: str = ""
    target_sheets: tuple[str, ...] = ()
    item_outcomes: list[ItemWriteOutcome] = field(default_factory=list)

    @property
    def operation(self) -> ExcelOperation:
        return ExcelOperation.PAYMENT_SYNC

    @property
    def sheet_name(self) -> str:
        return ", ".join(result.sheet_name for result in self.target_results.values())

    @property
    def sheet_created(self) -> bool:
        return any(result.sheet_created for result in self.target_results.values())

    @property
    def template_sheet_name(self) -> str | None:
        names = [
            result.template_sheet_name
            for result in self.target_results.values()
            if result.template_sheet_name
        ]
        return ", ".join(names) if names else None

    @property
    def inserted_rows(self) -> int:
        return sum(result.inserted_rows for result in self.target_results.values())

    @property
    def updated_rows(self) -> int:
        return sum(result.updated_rows for result in self.target_results.values())

    @property
    def unchanged_rows(self) -> int:
        return sum(result.unchanged_rows for result in self.target_results.values())

    @property
    def carrier_written_cells(self) -> int:
        return sum(
            result.carrier_written_cells for result in self.target_results.values()
        )

    @property
    def carrier_appended_cells(self) -> int:
        return sum(
            result.carrier_appended_cells for result in self.target_results.values()
        )

    @property
    def carrier_kept_cells(self) -> int:
        return sum(
            result.carrier_kept_cells for result in self.target_results.values()
        )

    @property
    def unmapped_carrier_items(self) -> int:
        return sum(
            result.unmapped_carrier_items
            for result in self.target_results.values()
        )

    @property
    def skipped_rows(self) -> int:
        return sum(result.skipped_rows for result in self.target_results.values())

    @property
    def conflict_count(self) -> int:
        return sum(result.conflict_count for result in self.target_results.values())

    @property
    def invoice_written_cells(self) -> int:
        return sum(
            result.invoice_written_cells for result in self.target_results.values()
        )


Plan = SyncPlan | PostingPlan | PaymentSyncPlan | PaymentSyncBatchPlan
Result = SyncResult | PostingResult | PaymentSyncResult
Resolution = SyncResolution | PostingResolution


def resolution_map(
    resolutions: Mapping[str, Any] | Sequence[Any] | None,
) -> dict[str, Any]:
    if resolutions is None:
        return {}
    if isinstance(resolutions, Mapping):
        return dict(resolutions)
    result: dict[str, Any] = {}
    for value in resolutions:
        conflict_id = getattr(value, "conflict_id", getattr(value, "id", None))
        if conflict_id is None:
            raise ValueError("Mỗi resolution phải có conflict_id.")
        result[str(conflict_id)] = value
    return result
