"""Durable, compatibility-checked Excel conflict choices."""

from __future__ import annotations

import hashlib
import json
from collections.abc import Mapping, Sequence
from dataclasses import asdict, dataclass, is_dataclass
from datetime import date, datetime
from enum import Enum
from pathlib import Path
from typing import Any

from app.repositories.excel_draft_repository import ExcelDraftRepository


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


def _code(value: Any) -> str:
    raw = getattr(value, "value", value)
    return str(raw or "").strip()


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


def _jsonable(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if is_dataclass(value) and not isinstance(value, type):
        return _jsonable(asdict(value))
    if isinstance(value, Mapping):
        return {
            str(key): _jsonable(item)
            for key, item in sorted(value.items(), key=lambda pair: str(pair[0]))
        }
    if isinstance(value, Sequence) and not isinstance(
        value, (str, bytes, bytearray)
    ):
        return [_jsonable(item) for item in value]
    to_dict = getattr(value, "to_dict", None)
    if callable(to_dict):
        return _jsonable(to_dict())
    return str(value)


def _digest(value: Any) -> str:
    payload = json.dumps(
        _jsonable(value),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _path_key(value: Any) -> str | None:
    if value in (None, ""):
        return None
    return str(Path(str(value)).expanduser().resolve(strict=False)).casefold()


def _fingerprint_key(value: Any) -> str | None:
    if value is None:
        return None
    raw = _value(value, "sha256", default=None)
    if raw not in (None, ""):
        return str(raw)
    if isinstance(value, Mapping):
        raw = value.get("sha256")
        if raw not in (None, ""):
            return str(raw)
    return _digest(value)


@dataclass(frozen=True, slots=True)
class ExcelDraftRestore:
    source_file_key: str
    resolutions: dict[str, Any]
    found: bool
    status: str | None
    updated_at: str | None
    current_conflict_count: int
    saved_choice_count: int
    restored_count: int
    target_compatible: bool
    last_error: str | None


class ExcelDraftService:
    """Keep the latest decisions per source file and replay compatible choices."""

    PAYLOAD_VERSION = 1

    def __init__(self, repository: ExcelDraftRepository) -> None:
        self.repository = repository

    def restore(self, plan: Any, operation: str) -> tuple[str, dict[str, Any]]:
        result = self.restore_with_info(plan, operation)
        return result.source_file_key, dict(result.resolutions)

    def restore_with_info(self, plan: Any, operation: str) -> ExcelDraftRestore:
        context = self._context(plan, operation)
        source_file_key = self._source_file_key(context)
        conflicts = self._conflicts(plan)
        record = self.repository.get_latest(source_file_key)
        if record is None:
            return ExcelDraftRestore(
                source_file_key=source_file_key,
                resolutions={},
                found=False,
                status=None,
                updated_at=None,
                current_conflict_count=len(conflicts),
                saved_choice_count=0,
                restored_count=0,
                target_compatible=True,
                last_error=None,
            )

        payload = record.payload
        entries = payload.get("entries", {})
        restored: dict[str, Any] = {}
        target_compatible = self._target_is_compatible(record.context, context)
        if target_compatible:
            for conflict in conflicts:
                conflict_id = self._conflict_id(conflict)
                entry = entries.get(conflict_id)
                if not isinstance(entry, Mapping):
                    continue
                if entry.get("signature") != self._conflict_signature(conflict):
                    continue
                resolution = entry.get("resolution")
                if isinstance(resolution, Mapping) and self._resolution_is_valid(
                    conflict, resolution
                ):
                    restored[conflict_id] = dict(resolution)

        globals_value = payload.get("globals", {})
        if (
            target_compatible
            and isinstance(globals_value, Mapping)
            and "selected_new_rows" in globals_value
        ):
            selected = globals_value.get("selected_new_rows")
            if isinstance(selected, list):
                valid_ids = {
                    str(_value(item, "item_id", default=""))
                    for item in _sequence(_value(plan, "new_rows", default=()))
                }
                restored["selected_new_rows"] = [
                    str(item_id)
                    for item_id in selected
                    if str(item_id) in valid_ids
                ]
        if target_compatible and isinstance(globals_value, Mapping):
            saved_splits = globals_value.get("split_document_ids")
            current_document_ids = {
                str(_value(group, "source_document_id", default=""))
                for group in _sequence(_value(plan, "source_groups", default=()))
                if str(_value(group, "source_document_id", default=""))
            }
            if isinstance(saved_splits, list):
                compatible_splits = sorted(
                    str(document_id)
                    for document_id in saved_splits
                    if str(document_id) in current_document_ids
                )
                if compatible_splits:
                    restored["split_document_ids"] = compatible_splits
            saved_assignments = globals_value.get("group_target_sheets")
            groups = {
                str(_value(group, "group_id", default="")): group
                for group in _sequence(_value(plan, "source_groups", default=()))
                if str(_value(group, "group_id", default=""))
            }
            candidate_names = {
                str(_value(candidate, "sheet_name", "name", default="")).strip()
                for candidate in _sequence(
                    _value(plan, "sheet_candidates", "target_sheet_candidates", default=())
                )
                if str(_value(candidate, "sheet_name", "name", default="")).strip()
            }
            if isinstance(saved_assignments, Mapping) and groups:
                compatible_assignments = {
                    str(group_id): str(sheet_name)
                    for group_id, sheet_name in saved_assignments.items()
                    if str(group_id) in groups
                    and str(sheet_name).strip()
                    and (not candidate_names or str(sheet_name) in candidate_names)
                }
                if compatible_assignments:
                    restored["group_target_sheets"] = compatible_assignments
        restored_count = sum(
            key not in {
                "selected_new_rows", "group_target_sheets", "split_document_ids"
            }
            for key in restored
        )
        return ExcelDraftRestore(
            source_file_key=source_file_key,
            resolutions=restored,
            found=True,
            status=record.status,
            updated_at=record.updated_at,
            current_conflict_count=len(conflicts),
            saved_choice_count=len(entries),
            restored_count=restored_count,
            target_compatible=target_compatible,
            last_error=record.last_error,
        )

    def save(
        self,
        plan: Any,
        operation: str,
        resolutions: Mapping[str, Any] | Sequence[Any] | None,
    ) -> str:
        context = self._context(plan, operation)
        source_file_key = self._source_file_key(context)
        run_id = self._run_id(context.get("run_id"))
        previous = self.repository.get_latest(source_file_key)
        payload: dict[str, Any] = (
            dict(previous.payload) if previous is not None else {}
        )
        entries = dict(payload.get("entries", {}))
        globals_value = dict(payload.get("globals", {}))
        group_assignments = {
            str(_value(group, "group_id")): str(_value(group, "target_sheet"))
            for group in _sequence(_value(plan, "source_groups", default=()))
            if _value(group, "group_id", default=None) not in (None, "")
            and _value(group, "target_sheet", default=None) not in (None, "")
        }
        if group_assignments:
            globals_value["group_target_sheets"] = group_assignments
        split_document_ids = sorted(
            str(value)
            for value in _sequence(_value(plan, "split_document_ids", default=()))
            if str(value).strip()
        )
        if split_document_ids:
            globals_value["split_document_ids"] = split_document_ids
        conflicts = {
            self._conflict_id(conflict): conflict for conflict in self._conflicts(plan)
        }
        if resolutions is None:
            resolution_values: dict[str, Any] = {}
        elif isinstance(resolutions, Mapping):
            resolution_values = dict(resolutions)
        else:
            resolution_values = {
                str(_value(item, "conflict_id", "id")): item
                for item in resolutions
                if _value(item, "conflict_id", "id", default=None) is not None
            }
        for key, raw_resolution in resolution_values.items():
            key = str(key)
            if key == "selected_new_rows":
                globals_value[key] = [
                    str(item_id) for item_id in _sequence(raw_resolution)
                ]
                continue
            if key == "group_target_sheets" and isinstance(raw_resolution, Mapping):
                globals_value[key] = {
                    str(group_id): str(sheet_name)
                    for group_id, sheet_name in raw_resolution.items()
                    if str(group_id).strip() and str(sheet_name).strip()
                }
                continue
            if key == "split_document_ids":
                globals_value[key] = sorted(
                    str(value) for value in _sequence(raw_resolution) if str(value).strip()
                )
                continue
            conflict = conflicts.get(key)
            if conflict is None:
                continue
            resolution = _jsonable(raw_resolution)
            if not isinstance(resolution, Mapping):
                continue
            entries[key] = {
                "signature": self._conflict_signature(conflict),
                "resolution": dict(resolution),
            }
        payload = {
            "version": self.PAYLOAD_VERSION,
            "entries": entries,
            "globals": globals_value,
        }
        self.repository.save(
            source_file_key=source_file_key,
            operation=operation,
            context=context,
            payload=payload,
            run_id=run_id,
        )
        return source_file_key

    def mark_failed(self, context_key: str | None, error: object) -> None:
        if context_key:
            self.repository.mark_failed(context_key, error)

    def mark_completed(self, context_key: str | None) -> None:
        if context_key:
            self.repository.mark_succeeded(context_key)

    def mark_cancelled(self, context_key: str | None) -> None:
        if context_key:
            self.repository.mark_cancelled(context_key)

    def _context(self, plan: Any, operation: str) -> dict[str, Any]:
        source_path = _value(plan, "source_path", "batch_path", default=None)
        source_fingerprint = _value(
            plan, "source_fingerprint", "source_snapshot_fingerprint", default=None
        )
        source_identity = _value(plan, "batch_hash", default=None)
        if source_identity in (None, ""):
            source_identity = _fingerprint_key(source_fingerprint)

        source_sheet = _value(plan, "source_sheet", default=None)
        if source_sheet in (None, ""):
            candidates = _sequence(_value(plan, "month_candidates", default=()))
            source_sheets = sorted(
                {
                    str(_value(candidate, "source_sheet", default="")).strip()
                    for candidate in candidates
                    if str(_value(candidate, "source_sheet", default="")).strip()
                }
            )
        else:
            source_sheets = [str(source_sheet)]
        return {
            "operation": operation,
            "source_path": _path_key(source_path),
            "source_identity": source_identity,
            "target_path": _path_key(_value(plan, "target_path", default=None)),
            "source_sheets": source_sheets,
            "selected_sheet": _value(
                plan, "selected_sheet", "selected_target_sheet", default=None
            ),
            "batch_id": _value(plan, "batch_id", default=None),
            "run_id": _value(plan, "run_id", default=None),
            "source_kind": _code(_value(plan, "source_kind", default="")),
        }

    @staticmethod
    def _source_file_key(context: Mapping[str, Any]) -> str:
        operation = str(context.get("operation") or "")
        batch_id = context.get("batch_id")
        if batch_id not in (None, ""):
            return f"{operation}:batch:{batch_id}"
        source_path = context.get("source_path")
        if source_path not in (None, ""):
            return f"{operation}:path:{str(source_path).casefold()}"
        source_identity = context.get("source_identity")
        if source_identity not in (None, ""):
            return f"{operation}:identity:{source_identity}"
        return f"{operation}:anonymous:{_digest(context)}"

    @staticmethod
    def _run_id(value: Any) -> int | None:
        if value in (None, ""):
            return None
        try:
            return int(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _target_is_compatible(
        previous_context: Mapping[str, Any], current_context: Mapping[str, Any]
    ) -> bool:
        previous = previous_context.get("target_path")
        current = current_context.get("target_path")
        if previous in (None, "") or current in (None, ""):
            return True
        return str(previous).casefold() == str(current).casefold()

    @staticmethod
    def _conflicts(plan: Any) -> tuple[Any, ...]:
        return _sequence(_value(plan, "conflicts", default=()))

    @staticmethod
    def _conflict_id(conflict: Any) -> str:
        return str(_value(conflict, "conflict_id", "id", default=""))

    def _conflict_signature(self, conflict: Any) -> str:
        candidates = []
        for candidate in _sequence(
            _value(conflict, "row_candidates", "candidates", default=())
        ):
            candidates.append(
                {
                    "sheet": _value(candidate, "source_sheet", "sheet_name"),
                    "row": _value(candidate, "row", "target_row"),
                    "sqt": _value(candidate, "sqt"),
                    "container": _value(candidate, "container"),
                    "vessel": _value(candidate, "vessel"),
                }
            )
        return _digest(
            {
                "type": _code(_value(conflict, "conflict_type", "type")),
                "container": _value(conflict, "container"),
                "bl": _value(conflict, "bl"),
                "sqt": _value(conflict, "sqt"),
                "fee": _value(conflict, "fee"),
                "amount": _value(conflict, "amount"),
                "sheet": _value(conflict, "sheet_name"),
                "row": _value(conflict, "target_row"),
                "column": _value(conflict, "target_column"),
                "cell": _value(conflict, "target_cell"),
                "current": _value(conflict, "current_value"),
                "allowed": [
                    _code(action)
                    for action in _sequence(
                        _value(conflict, "allowed_actions", "options", default=())
                    )
                ],
                "candidates": candidates,
                "details": _value(conflict, "details", default={}),
            }
        )

    @staticmethod
    def _resolution_is_valid(conflict: Any, resolution: Mapping[str, Any]) -> bool:
        action = _code(resolution.get("action"))
        if not action:
            return False
        if action in {"CANCEL", "CANCEL_ALL"}:
            return False
        allowed = {
            _code(value)
            for value in _sequence(
                _value(conflict, "allowed_actions", "options", default=())
            )
        }
        if allowed and action not in allowed:
            return False

        details = _value(conflict, "details", default={}) or {}
        if action == "SELECT_ROW":
            selected_row = resolution.get("selected_row", resolution.get("row"))
            selected_sheet = resolution.get(
                "selected_source_sheet", resolution.get("selected_sheet")
            )
            if selected_row is None:
                return False
            for candidate in _sequence(
                _value(conflict, "row_candidates", "candidates", default=())
            ):
                row = _value(candidate, "row", "target_row")
                sheet = _value(candidate, "source_sheet", "sheet_name")
                if row is None or int(row) != int(selected_row):
                    continue
                if selected_sheet in (None, "") or sheet in (None, ""):
                    return True
                if str(sheet) == str(selected_sheet):
                    return True
            return False
        if action == "SELECT_INVOICE":
            selected = str(resolution.get("selected_invoice") or "").casefold()
            return selected in {
                str(value).casefold()
                for value in _sequence(details.get("invoice_candidates", ()))
            }
        if action == "SELECT_CARRIER":
            selected = str(resolution.get("selected_carrier") or "").casefold()
            candidates = details.get(
                "carrier_candidates", details.get("existing_carrier_candidates", ())
            )
            return selected in {
                str(value).casefold() for value in _sequence(candidates)
            }
        if action == "SELECT_SOURCE_ITEM":
            selected = resolution.get("selected_source_item_index")
            return selected is not None and int(selected) in {
                int(_value(option, "source_item_index"))
                for option in _sequence(details.get("source_item_options", ()))
            }
        if action == "SELECT_FEE":
            return bool(resolution.get("selected_fee", resolution.get("fee")))
        return True


__all__ = ["ExcelDraftRestore", "ExcelDraftService"]
