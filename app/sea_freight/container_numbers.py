"""Chuẩn hóa và kiểm tra số container theo ISO 6346."""

from __future__ import annotations

import re

_CONTAINER_RE = re.compile(r"^[A-Z]{4}[0-9]{7}$")
_ISO_LETTER_VALUES = {
    letter: value
    for letter, value in zip(
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
        (10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 23, 24,
         25, 26, 27, 28, 29, 30, 31, 32, 34, 35, 36, 37, 38),
        strict=True,
    )
}


class InvalidContainerNumber(ValueError):
    pass


def normalize_container_number(value: object) -> str:
    if not isinstance(value, str):
        return ""
    return re.sub(r"[^A-Z0-9]", "", value.upper())


def iso6346_check_digit(first_ten_characters: str) -> int:
    normalized = normalize_container_number(first_ten_characters)
    if re.fullmatch(r"[A-Z]{4}[0-9]{6}", normalized) is None:
        raise InvalidContainerNumber(
            "Phần thân container phải có 4 chữ cái và 6 chữ số."
        )
    total = sum(
        (_ISO_LETTER_VALUES[character] if character.isalpha() else int(character))
        * (2**position)
        for position, character in enumerate(normalized)
    )
    return (total % 11) % 10


def validate_iso6346(value: object) -> str:
    normalized = normalize_container_number(value)
    if _CONTAINER_RE.fullmatch(normalized) is None:
        raise InvalidContainerNumber(
            f"Container {value!r} không đúng mẫu 4 chữ cái và 7 chữ số."
        )
    expected = iso6346_check_digit(normalized[:10])
    if int(normalized[-1]) != expected:
        raise InvalidContainerNumber(
            f"Container {normalized} sai số kiểm tra ISO 6346 (đúng phải là {expected})."
        )
    return normalized


def allocate_integer_amount(total: int, count: int) -> tuple[int, ...]:
    if type(total) is not int or total < 0:
        raise ValueError("Tổng tiền phải là số nguyên không âm.")
    if type(count) is not int or count <= 0:
        raise ValueError("Số container phải là số nguyên dương.")
    base, remainder = divmod(total, count)
    # Người dùng kiểm tra các dòng đầu với cùng một mức tiền; toàn bộ phần lẻ
    # được dồn vào dòng cuối để tổng luôn khớp tuyệt đối với hóa đơn.
    return tuple(
        base + (remainder if index == count - 1 else 0)
        for index in range(count)
    )


__all__ = [
    "InvalidContainerNumber",
    "allocate_integer_amount",
    "iso6346_check_digit",
    "normalize_container_number",
    "validate_iso6346",
]
