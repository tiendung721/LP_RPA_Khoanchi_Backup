"""Shared contracts for recoverable Excel conflict review."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from typing import Any


def _value(source: Any, *names: str, default: Any = None) -> Any:
    if source is None:
        return default
    if isinstance(source, Mapping):
        for name in names:
            if name in source:
                return source[name]
        return default
    for name in names:
        if hasattr(source, name):
            return getattr(source, name)
    return default


def _sequence(value: Any) -> tuple[Any, ...]:
    if value is None:
        return ()
    if isinstance(value, Mapping):
        return tuple(value.values())
    if isinstance(value, Sequence) and not isinstance(
        value, (str, bytes, bytearray)
    ):
        return tuple(value)
    try:
        return tuple(value)
    except TypeError:
        return (value,)


@dataclass(frozen=True, slots=True)
class CorrectionIssue:
    """One user-correctable problem linked to conflict-table rows."""

    issue_id: str
    message: str
    conflict_ids: tuple[str, ...] = ()
    item_index: int | None = None
    item_id: str | None = None
    source_row: int | None = None
    field: str | None = None
    sheet_name: str | None = None
    target_row: int | None = None
    target_column: int | str | None = None
    target_cell: str | None = None

    @classmethod
    def from_conflict(
        cls,
        conflict: Any,
        *,
        message: str | None = None,
        issue_id: str | None = None,
        field: str | None = None,
    ) -> "CorrectionIssue":
        conflict_id = str(_value(conflict, "conflict_id", "id", default=""))
        details = _value(conflict, "details", default={}) or {}
        raw_item_index = _value(conflict, "item_index", default=None)
        raw_source_row = _value(conflict, "source_row", default=None)
        return cls(
            issue_id=issue_id or f"resolution:{conflict_id}",
            message=str(
                message
                or _value(
                    conflict,
                    "message",
                    "reason",
                    default="Lựa chọn chưa hợp lệ.",
                )
            ),
            conflict_ids=(conflict_id,) if conflict_id else (),
            item_index=(
                int(raw_item_index) if raw_item_index is not None else None
            ),
            item_id=(
                str(_value(conflict, "item_id"))
                if _value(conflict, "item_id", default=None) not in (None, "")
                else None
            ),
            source_row=(
                int(raw_source_row) if raw_source_row is not None else None
            ),
            field=field or _value(details, "field", default=None),
            sheet_name=_value(
                conflict, "sheet_name", "sheet", "target_sheet", default=None
            ),
            target_row=_value(
                conflict, "target_row", "row", "row_number", default=None
            ),
            target_column=_value(
                conflict, "target_column", "column", default=None
            ),
            target_cell=_value(
                conflict, "target_cell", "cell", "cell_address", default=None
            ),
        )


@dataclass(slots=True)
class ReviewOutcome:
    """Read-only result of replaying conflict choices from a base plan."""

    prepared_plan: Any | None
    conflicts: tuple[Any, ...] = ()
    issues: tuple[CorrectionIssue, ...] = ()
    resolutions: dict[str, Any] = field(default_factory=dict)

    @property
    def ready(self) -> bool:
        return self.prepared_plan is not None and not self.issues

    @classmethod
    def needs_correction(
        cls,
        *,
        conflicts: Sequence[Any],
        issues: Sequence[CorrectionIssue] | None = None,
        resolutions: Mapping[str, Any] | None = None,
    ) -> "ReviewOutcome":
        visible = tuple(conflicts)
        return cls(
            prepared_plan=None,
            conflicts=visible,
            issues=tuple(
                issues
                if issues is not None
                else (CorrectionIssue.from_conflict(item) for item in visible)
            ),
            resolutions=dict(resolutions or {}),
        )


class CorrectionRequiredError(RuntimeError):
    """Raised before writing when choices can be corrected in the review UI."""

    def __init__(
        self,
        issues: Sequence[CorrectionIssue],
        message: str = "Dữ liệu xung đột chưa hợp lệ.",
    ) -> None:
        self.issues = tuple(issues)
        super().__init__(message)


class SourceDataChangedError(RuntimeError):
    """The non-workbook source changed and the base plan must be rebuilt."""

    def __init__(self, label: str, message: str | None = None) -> None:
        self.label = label
        super().__init__(
            message or f"{label} đã thay đổi sau khi phân tích; vui lòng đọc lại."
        )


def issues_for_conflicts(conflicts: Sequence[Any]) -> tuple[CorrectionIssue, ...]:
    return tuple(CorrectionIssue.from_conflict(conflict) for conflict in conflicts)


def _code(value: Any) -> str:
    return str(getattr(value, "value", value) or "").strip()


def _resolution_value(value: Any, *names: str) -> Any:
    return _value(value, *names, default=None)


def validate_conflict_resolutions(
    conflicts: Sequence[Any],
    resolutions: Mapping[str, Any] | None,
) -> tuple[CorrectionIssue, ...]:
    """Collect every invalid editable choice instead of failing at the first row."""

    resolved = dict(resolutions or {})
    issues: list[CorrectionIssue] = []
    for conflict in conflicts:
        conflict_id = str(_value(conflict, "conflict_id", "id", default=""))
        resolution = resolved.get(conflict_id)
        action = _code(_resolution_value(resolution, "action"))
        if not action:
            issues.append(
                CorrectionIssue.from_conflict(
                    conflict,
                    message="Chưa chọn cách xử lý.",
                    issue_id=f"missing-action:{conflict_id}",
                    field="action",
                )
            )
            continue

        allowed = {
            _code(item)
            for item in _sequence(
                _value(
                    conflict,
                    "allowed_actions",
                    "allowed_resolutions",
                    "actions",
                    "options",
                    default=(),
                )
            )
        }
        if allowed and action not in allowed:
            issues.append(
                CorrectionIssue.from_conflict(
                    conflict,
                    message="Cách xử lý đã chọn không còn hợp lệ.",
                    issue_id=f"invalid-action:{conflict_id}",
                    field="action",
                )
            )
            continue

        details = _value(conflict, "details", default={}) or {}
        message: str | None = None
        field: str | None = None
        if action == "SELECT_ROW":
            selected_row = _resolution_value(resolution, "selected_row", "row")
            candidates = _sequence(
                _value(
                    conflict,
                    "row_candidates",
                    "candidate_rows",
                    "candidates",
                    default=(),
                )
            )
            valid_rows = {
                int(row)
                for candidate in candidates
                if (row := _value(candidate, "row", "target_row", default=None))
                is not None
            }
            if selected_row is None:
                message, field = "Chưa chọn dòng đích.", "selected_row"
            elif valid_rows and int(selected_row) not in valid_rows:
                message, field = (
                    "Dòng đích đã chọn không còn hợp lệ.",
                    "selected_row",
                )
            else:
                selected_sheet = _resolution_value(
                    resolution,
                    "selected_source_sheet",
                    "selected_sheet",
                    "selected_sheet_name",
                )
                valid_pairs = {
                    (
                        str(
                            _value(
                                candidate,
                                "source_sheet",
                                "sheet_name",
                                default="",
                            )
                            or ""
                        ),
                        int(_value(candidate, "row", "target_row")),
                    )
                    for candidate in candidates
                    if _value(candidate, "row", "target_row", default=None)
                    is not None
                }
                if (
                    selected_sheet not in (None, "")
                    and valid_pairs
                    and (str(selected_sheet), int(selected_row)) not in valid_pairs
                ):
                    message, field = (
                        "Sheet và dòng đích đã chọn không còn tương thích.",
                        "selected_row",
                    )
        elif action == "SELECT_FEE":
            selected = _resolution_value(resolution, "selected_fee", "fee")
            candidates = {
                _code(item)
                for item in _sequence(
                    _value(
                        conflict,
                        "valid_fee_codes",
                        "fee_candidates",
                        default=(),
                    )
                )
            }
            if selected in (None, ""):
                message, field = "Chưa chọn mã phí.", "selected_fee"
            elif candidates and _code(selected) not in candidates:
                message, field = "Mã phí đã chọn không còn hợp lệ.", "selected_fee"
        elif action == "SELECT_SHEET":
            selected = _resolution_value(
                resolution, "selected_sheet_name", "selected_sheet", "sheet_name"
            )
            raw_candidates = _sequence(
                _value(
                    conflict,
                    "sheet_candidates",
                    "target_sheet_candidates",
                    default=_value(
                        details,
                        "sheet_candidates",
                        "target_sheet_candidates",
                        default=(),
                    ),
                )
            )
            candidates = {
                str(
                    _value(
                        item,
                        "target_sheet",
                        "sheet_name",
                        "source_sheet",
                        default=item if isinstance(item, str) else "",
                    )
                )
                for item in raw_candidates
            }
            candidates.discard("")
            if selected in (None, ""):
                message, field = "Chưa chọn sheet.", "selected_sheet"
            elif candidates and str(selected) not in candidates:
                message, field = "Sheet đã chọn không còn hợp lệ.", "selected_sheet"
        elif action == "SELECT_MONTH":
            selected = _resolution_value(resolution, "selected_month", "month")
            raw_candidates = _sequence(
                _value(
                    conflict,
                    "month_candidates",
                    "sheet_candidates",
                    default=_value(
                        details,
                        "month_candidates",
                        "sheet_candidates",
                        default=(),
                    ),
                )
            )
            candidates = {
                int(month)
                for item in raw_candidates
                if (month := _value(item, "month", "month_number", default=None))
                is not None
            }
            if selected is None:
                message, field = "Chưa chọn tháng.", "selected_month"
            elif candidates and int(selected) not in candidates:
                message, field = "Tháng đã chọn không còn hợp lệ.", "selected_month"
        elif action == "SELECT_INVOICE":
            selected = _resolution_value(resolution, "selected_invoice")
            candidates = {
                str(item).strip().casefold()
                for item in _sequence(
                    _value(details, "invoice_candidates", default=())
                )
            }
            if selected in (None, ""):
                message, field = "Chưa chọn Số HĐ.", "selected_invoice"
            elif candidates and str(selected).strip().casefold() not in candidates:
                message, field = (
                    "Số HĐ đã chọn không còn hợp lệ.",
                    "selected_invoice",
                )
        elif action == "SELECT_SOURCE_ITEM":
            selected = _resolution_value(resolution, "selected_source_item_index")
            candidates = {
                int(_value(item, "source_item_index"))
                for item in _sequence(
                    _value(details, "source_item_options", default=())
                )
                if _value(item, "source_item_index", default=None) is not None
            }
            if selected is None:
                message, field = "Chưa chọn dòng JSON.", "selected_source_item_index"
            elif candidates and int(selected) not in candidates:
                message, field = (
                    "Dòng JSON đã chọn không còn hợp lệ.",
                    "selected_source_item_index",
                )
        elif action == "SELECT_CARRIER" or (
            action == "KEEP_EXISTING"
            and bool(
                _value(
                    details,
                    "requires_existing_carrier_selection",
                    default=False,
                )
            )
        ):
            selected = _resolution_value(resolution, "selected_carrier")
            candidate_field = (
                "existing_carrier_candidates"
                if action == "KEEP_EXISTING"
                else "carrier_candidates"
            )
            candidates = {
                str(item).strip().casefold()
                for item in _sequence(_value(details, candidate_field, default=()))
            }
            if selected in (None, ""):
                message, field = "Chưa chọn bên vận tải.", "selected_carrier"
            elif candidates and str(selected).strip().casefold() not in candidates:
                message, field = (
                    "Bên vận tải đã chọn không còn hợp lệ.",
                    "selected_carrier",
                )

        if message:
            issues.append(
                CorrectionIssue.from_conflict(
                    conflict,
                    message=message,
                    issue_id=f"invalid-choice:{conflict_id}:{field}",
                    field=field,
                )
            )
    return tuple(issues)


__all__ = [
    "CorrectionIssue",
    "CorrectionRequiredError",
    "ReviewOutcome",
    "SourceDataChangedError",
    "issues_for_conflicts",
    "validate_conflict_resolutions",
]
