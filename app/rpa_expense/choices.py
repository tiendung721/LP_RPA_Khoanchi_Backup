"""Ghi nhớ lựa chọn RPA gần nhất theo file BK và sheet."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Sequence

from app.repositories.excel_draft_repository import ExcelDraftRepository


def _path_key(path: str | Path) -> str:
    return str(Path(path).expanduser().resolve(strict=False)).casefold()


def _item_signature(item: Any) -> str:
    amounts = getattr(item, "amounts", None)
    values = amounts.to_dict() if callable(getattr(amounts, "to_dict", None)) else {}
    raw = json.dumps(
        {"sqt": str(getattr(item, "sqt", "")), "amounts": values},
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


@dataclass(frozen=True, slots=True)
class RpaChoiceRestore:
    selected_sqt: tuple[str, ...]
    found: bool = False
    saved_count: int = 0
    restored_count: int = 0
    skipped_sqt: tuple[str, ...] = ()
    updated_at: str | None = None
    status: str | None = None


class RpaChoiceService:
    OPERATION = "rpa_expense"
    SHEET_OPERATION = "rpa_expense_sheet"
    PAYLOAD_VERSION = 2

    def __init__(self, repository: ExcelDraftRepository) -> None:
        self.repository = repository

    def save_sheet(self, bk_path: str | Path, sheet_name: str) -> str:
        key = self._sheet_key(bk_path)
        self.repository.save(
            source_file_key=key,
            operation=self.SHEET_OPERATION,
            context={"bk_path": _path_key(bk_path)},
            payload={"version": self.PAYLOAD_VERSION, "selected_sheet": sheet_name},
        )
        return key

    def restore_sheet(
        self, bk_path: str | Path, candidates: Sequence[Any]
    ) -> str | None:
        record = self.repository.get_latest(self._sheet_key(bk_path))
        if record is None:
            return None
        selected = str(record.payload.get("selected_sheet") or "")
        available = {
            str(getattr(candidate, "sheet_name", getattr(candidate, "target_sheet", "")))
            for candidate in candidates
        }
        return selected if selected and selected in available else None

    def save_selection(self, plan: Any, selected_sqt: Iterable[str]) -> str:
        selected = tuple(dict.fromkeys(str(value).strip() for value in selected_sqt))
        lookup = {
            str(getattr(item, "sqt", "")): item
            for item in getattr(plan, "items", ())
        }
        selected = tuple(value for value in selected if value in lookup)
        key = self._selection_key(plan.bk_path, plan.sheet_name)
        self.repository.save(
            source_file_key=key,
            operation=self.OPERATION,
            context={
                "bk_path": _path_key(plan.bk_path),
                "sheet_name": str(plan.sheet_name),
                "source_fingerprint": getattr(plan.fingerprint, "sha256", None),
            },
            payload={
                "version": self.PAYLOAD_VERSION,
                "selected_sqt": list(selected),
                "signatures": {
                    value: _item_signature(lookup[value]) for value in selected
                },
            },
        )
        return key

    def restore_selection(self, plan: Any) -> RpaChoiceRestore:
        key = self._selection_key(plan.bk_path, plan.sheet_name)
        record = self.repository.get_latest(key)
        if record is None:
            return RpaChoiceRestore(())
        saved = tuple(
            str(value)
            for value in record.payload.get("selected_sqt", ())
            if str(value).strip()
        )
        signatures = record.payload.get("signatures", {})
        lookup = {
            str(getattr(item, "sqt", "")): item
            for item in getattr(plan, "items", ())
        }
        restored: list[str] = []
        skipped: list[str] = []
        for sqt in saved:
            item = lookup.get(sqt)
            if (
                item is None
                or not bool(getattr(item, "can_run", False))
                or (
                    isinstance(signatures, dict)
                    and sqt in signatures
                    and signatures[sqt] != _item_signature(item)
                )
            ):
                skipped.append(sqt)
                continue
            restored.append(sqt)
        return RpaChoiceRestore(
            selected_sqt=tuple(restored),
            found=True,
            saved_count=len(saved),
            restored_count=len(restored),
            skipped_sqt=tuple(skipped),
            updated_at=record.updated_at,
            status=record.status,
        )

    def mark_failed(self, plan: Any, error: object) -> None:
        self.repository.mark_failed(
            self._selection_key(plan.bk_path, plan.sheet_name), error
        )

    def mark_launched(self, plan: Any) -> None:
        self.repository.mark_succeeded(
            self._selection_key(plan.bk_path, plan.sheet_name)
        )

    def clear_selection(self, plan: Any) -> bool:
        return self.repository.delete(
            self._selection_key(plan.bk_path, plan.sheet_name)
        )

    def clear_sheet(self, bk_path: str | Path) -> bool:
        return self.repository.delete(self._sheet_key(bk_path))

    @classmethod
    def _sheet_key(cls, bk_path: str | Path) -> str:
        return f"{cls.SHEET_OPERATION}:path:{_path_key(bk_path)}"

    @classmethod
    def _selection_key(cls, bk_path: str | Path, sheet_name: str) -> str:
        return (
            f"{cls.OPERATION}:path:{_path_key(bk_path)}:"
            f"sheet:{str(sheet_name).strip().casefold()}"
        )


__all__ = ["RpaChoiceRestore", "RpaChoiceService"]
