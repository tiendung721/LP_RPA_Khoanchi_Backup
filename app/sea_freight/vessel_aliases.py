"""Chuẩn hóa và so khớp alias tàu/chuyến theo các quy tắc kiểm soát."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass


_NON_ALNUM = re.compile(r"[^A-Z0-9]+")
_VOYAGE_KEY = re.compile(r"^(?P<prefix>[A-Z]*)(?P<number>\d+)(?P<suffix>[A-Z]*)$")

# Chỉ bổ sung các cặp đã được xác nhận từ dữ liệu thực tế. Không dùng fuzzy
# matching ở đây vì hàm này có quyền tự động chọn dòng BK.
_VESSEL_NAME_ALIASES = {
    "VSICOPROMOTE": "VSCPROMOTE",
    "VSCPROMOTE": "VSCPROMOTE",
    "BIENDONGMARINE": "BIENDONGMARINER",
    "BIENDONGMARINER": "BIENDONGMARINER",
}


@dataclass(frozen=True, slots=True)
class _VoyageIdentity:
    prefix: str
    number: str
    suffix: str


def normalize_match_key(value: object) -> str:
    if not isinstance(value, str):
        return ""
    decomposed = unicodedata.normalize("NFKD", value.strip().upper())
    ascii_like = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
    return _NON_ALNUM.sub("", ascii_like)


def canonical_vessel_key(value: object) -> str:
    key = normalize_match_key(value)
    return _VESSEL_NAME_ALIASES.get(key, key)


def vessel_alias_keys(value: object) -> tuple[str, ...]:
    canonical = canonical_vessel_key(value)
    if not canonical:
        return ()
    return tuple(
        sorted(
            {
                canonical,
                *(
                    alias
                    for alias, target in _VESSEL_NAME_ALIASES.items()
                    if target == canonical
                ),
            },
            key=lambda item: (-len(item), item),
        )
    )


def _voyage_identity(value: object) -> _VoyageIdentity | None:
    key = normalize_match_key(value)
    match = _VOYAGE_KEY.fullmatch(key)
    if match is None:
        return None
    return _VoyageIdentity(
        prefix=match.group("prefix"),
        number=match.group("number"),
        suffix=match.group("suffix"),
    )


def vessel_voyage_alias_equivalent(
    left_vessel_name: object,
    left_voyage_no: object,
    right_vessel_name: object,
    right_voyage_no: object,
) -> bool:
    """Khớp exact hoặc chỉ khác việc một bên thiếu tiền tố chuyến."""

    if canonical_vessel_key(left_vessel_name) != canonical_vessel_key(right_vessel_name):
        return False
    left = _voyage_identity(left_voyage_no)
    right = _voyage_identity(right_voyage_no)
    if left is None or right is None:
        return False
    if (left.number, left.suffix) != (right.number, right.suffix):
        return False
    return left.prefix == right.prefix or not left.prefix or not right.prefix


def _split_bk_vessel_voyage(
    raw_value: object,
    expected_vessel_key: str,
) -> tuple[str, str, str] | None:
    if not isinstance(raw_value, str):
        return None
    display = " ".join(raw_value.strip().split())
    combined_key = normalize_match_key(display)
    if not expected_vessel_key or not combined_key.startswith(expected_vessel_key):
        return None
    voyage_key = combined_key[len(expected_vessel_key) :]
    if _voyage_identity(voyage_key) is None:
        return None

    separators = " \t\r\n/-–—|:"
    for index in range(1, len(display)):
        vessel = display[:index].rstrip(separators)
        voyage = display[index:].lstrip(separators)
        if (
            normalize_match_key(vessel) == expected_vessel_key
            and normalize_match_key(voyage) == voyage_key
        ):
            return vessel, voyage, combined_key
    return None


def vessel_voyage_candidate_equivalent(
    vessel_name: object,
    voyage_no: object,
    candidate_raw: object,
) -> bool:
    """So một tàu/chuyến có cấu trúc với chuỗi gộp trên một dòng BK."""

    if not isinstance(candidate_raw, str) or not candidate_raw.strip():
        return False
    for vessel_key in vessel_alias_keys(vessel_name):
        split = _split_bk_vessel_voyage(candidate_raw, vessel_key)
        if split is None:
            continue
        candidate_vessel, candidate_voyage, _combined = split
        if vessel_voyage_alias_equivalent(
            vessel_name,
            voyage_no,
            candidate_vessel,
            candidate_voyage,
        ):
            return True
    return False


__all__ = [
    "canonical_vessel_key",
    "normalize_match_key",
    "vessel_alias_keys",
    "vessel_voyage_alias_equivalent",
    "vessel_voyage_candidate_equivalent",
]
