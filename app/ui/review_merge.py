"""Tìm và chuẩn bị gộp các khoản chi trong màn hình review GPT."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from datetime import date
from typing import Any

from app.constants import FEE_CODES
from app.services.validation_service import normalize_container

from .review_table_model import ReviewRow


MERGE_FIELDS = (
    "invoice_no",
    "invoice_date",
    "sqt",
    "bl",
    "carrier",
    "rule",
    "vessel_voyage_raw",
    "vessel_name",
    "voyage_no",
    "invoice_container_count",
    "container_count_basis",
)

FIELD_LABELS = {
    "source_document": "Chứng từ nguồn",
    "invoice_no": "Số HĐ",
    "invoice_date": "Ngày HĐ",
    "sqt": "SQT nguồn",
    "bl": "B/L",
    "carrier": "Bên vận tải",
    "rule": "Quy tắc tiền",
    "vessel_voyage_raw": "Tàu/chuyến nguyên văn",
    "vessel_name": "Tên tàu",
    "voyage_no": "Số chuyến",
    "invoice_container_count": "SL cont HĐ",
    "container_count_basis": "Căn cứ SL cont",
}


def _text(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    normalized = " ".join(value.split())
    return normalized.casefold() if normalized else None


def _invoice_key(value: Any) -> str | None:
    # So sánh như văn bản; không chuyển sang số để giữ số 0 ở đầu.
    return _text(value)


def _date_key(value: Any) -> str | None:
    if not isinstance(value, str) or not value:
        return None
    try:
        return value if date.fromisoformat(value).isoformat() == value else None
    except ValueError:
        return None


def _container_key(value: Any) -> str | None:
    try:
        return normalize_container(value)
    except TypeError:
        return None


def rows_match(left: ReviewRow, right: ReviewRow) -> bool:
    """Khớp khi có container, fee và ít nhất một dấu hiệu HĐ chung."""

    container = _container_key(left.cont)
    if not container or container != _container_key(right.cont):
        return False
    if not isinstance(left.fee, str) or left.fee not in FEE_CODES:
        return False
    if left.fee != right.fee or left.fee == "CXD":
        return False
    invoice_left = _invoice_key(left.invoice_no)
    invoice_right = _invoice_key(right.invoice_no)
    date_left = _date_key(left.invoice_date)
    date_right = _date_key(right.invoice_date)
    return bool(
        (invoice_left and invoice_left == invoice_right)
        or (date_left and date_left == date_right)
    )


def candidate_groups(
    rows: Sequence[ReviewRow],
    *,
    origins: Mapping[str, Sequence[ReviewRow]] | None = None,
) -> list[tuple[int, ...]]:
    """Nhóm theo thứ tự gốc; từng cặp trong nhóm phải trực tiếp khớp nhau."""

    used: set[int] = set()
    groups: list[tuple[int, ...]] = []

    def directly_matches(left: ReviewRow, right: ReviewRow) -> bool:
        if not rows_match(left, right):
            return False
        if origins is None:
            return True
        left_origins = origins.get(left.runtime_id, (left,))
        right_origins = origins.get(right.runtime_id, (right,))
        return all(
            rows_match(original_left, original_right)
            for original_left in left_origins
            for original_right in right_origins
        )

    for first in range(len(rows)):
        if first in used:
            continue
        members = [first]
        for candidate in range(first + 1, len(rows)):
            if candidate in used:
                continue
            if all(directly_matches(rows[candidate], rows[index]) for index in members):
                members.append(candidate)
        if len(members) > 1:
            groups.append(tuple(members))
            used.update(members)
    return groups


def _field_key(field: str, value: Any) -> Any:
    if value in (None, "") or (field == "container_count_basis" and value == "UNKNOWN"):
        return None
    if field == "invoice_no":
        return _invoice_key(value)
    if field in {"bl", "carrier", "rule", "vessel_voyage_raw", "vessel_name", "voyage_no"}:
        return _text(value)
    return value


def _source_key(row: ReviewRow) -> tuple[Any, Any]:
    return row.source_document_id, row.source_document_name


def choices_for_field(rows: Sequence[ReviewRow], field: str) -> list[Any]:
    choices: list[Any] = []
    seen: set[Any] = set()
    for row in rows:
        value = _source_key(row) if field == "source_document" else getattr(row, field)
        key = value if field == "source_document" else _field_key(field, value)
        try:
            hash(key)
        except TypeError:
            key = repr(key)
        if key is None or key in seen:
            continue
        seen.add(key)
        choices.append(value)
    return choices


def conflicting_fields(rows: Sequence[ReviewRow]) -> tuple[str, ...]:
    fields = ["source_document", *MERGE_FIELDS]
    return tuple(field for field in fields if len(choices_for_field(rows, field)) > 1)


def is_exact_duplicate_group(rows: Sequence[ReviewRow]) -> bool:
    objects = [row.to_object() for row in rows]
    return any(left == right for index, left in enumerate(objects) for right in objects[index + 1 :])


def has_valid_amounts(rows: Sequence[ReviewRow]) -> bool:
    return all(type(row.amount) is int and row.amount >= 0 for row in rows)


def merged_row(
    rows: Sequence[ReviewRow], selections: Mapping[str, Any] | None = None
) -> ReviewRow:
    """Tạo kết quả một dòng, giữ runtime ID của dòng xuất hiện đầu tiên."""

    if len(rows) < 2 or not has_valid_amounts(rows):
        raise ValueError("Cần ít nhất hai dòng có số tiền hợp lệ để gộp.")
    selected = selections or {}
    payload = rows[0].to_object()
    for field in MERGE_FIELDS:
        if field in selected:
            payload[field] = selected[field]
            continue
        choices = choices_for_field(rows, field)
        if choices:
            payload[field] = choices[0]
    source_choices = choices_for_field(rows, "source_document")
    source = selected.get("source_document", source_choices[0])
    payload["source_document_id"], payload["source_document_name"] = source
    payload["amount"] = sum(row.amount for row in rows)
    if payload["fee"] == "CB":
        payload["rule"] = "HD"
    return ReviewRow.from_mapping(payload, runtime_id=rows[0].runtime_id)
