"""Ánh xạ cấu trúc JSON vị trí sang model nội bộ.

Module này chỉ xác nhận hình dạng tối thiểu cần thiết để có thể ánh xạ và chỉnh
sửa dữ liệu. Các kiểu trường và quy tắc nghiệp vụ được báo chi tiết bởi
``ValidationService`` để người dùng có thể sửa lỗi thay vì mất toàn bộ lô.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
from typing import Any

from app.constants import LEGACY_SCHEMA_VERSION, SCHEMA_VERSION
from app.models import BatchDocument, DataRow


class SchemaError(ValueError):
    """Dữ liệu không thể ánh xạ an toàn sang schema ``v``/``d``."""

    def __init__(
        self,
        message: str,
        *,
        code: str = "invalid_schema",
        row_index: int | None = None,
        field: str | None = None,
    ) -> None:
        super().__init__(message)
        self.code = code
        self.row_index = row_index
        self.field = field


def parse_document(raw: object) -> BatchDocument:
    """Parse một object Python, giữ nguyên giá trị và thứ tự dòng."""

    if not isinstance(raw, Mapping):
        raise SchemaError(
            "Đối tượng gốc của JSON phải là object.",
            code="root_not_object",
        )
    keys = set(raw.keys())
    expected = {"v", "d"}
    if keys != expected:
        missing = sorted(expected.difference(keys))
        extra = sorted(str(key) for key in keys.difference(expected))
        details: list[str] = []
        if missing:
            details.append(f"thiếu khóa {', '.join(missing)}")
        if extra:
            details.append(f"thừa khóa {', '.join(extra)}")
        suffix = f" ({'; '.join(details)})" if details else ""
        raise SchemaError(
            f"Đối tượng gốc phải có đúng hai khóa v và d{suffix}.",
            code="root_keys",
        )

    version = raw["v"]
    if type(version) is not int or version not in {
        LEGACY_SCHEMA_VERSION,
        SCHEMA_VERSION,
    }:
        raise SchemaError(
            "Khóa v phải là số nguyên 1 hoặc 2.",
            code="invalid_version",
            field="v",
        )

    data = raw["d"]
    if not isinstance(data, list):
        raise SchemaError(
            "Khóa d phải là một mảng.",
            code="data_not_array",
            field="d",
        )

    rows: list[DataRow] = []
    v2_keys = set(DataRow(None, None, "CXD", None, None).to_object())
    for index, item in enumerate(data):
        if version == LEGACY_SCHEMA_VERSION:
            if not isinstance(item, list):
                raise SchemaError(
                    f"Dòng {index + 1} của schema v1 phải là mảng.",
                    code="row_not_array",
                    row_index=index,
                )
            if len(item) != 7:
                raise SchemaError(
                    f"Dòng {index + 1} của schema v1 phải là mảng 7 phần tử.",
                    code="row_length",
                    row_index=index,
                )
            rows.append(DataRow.from_sequence(item))
            continue
        if not isinstance(item, Mapping):
            raise SchemaError(
                f"Dòng {index + 1} của schema v2 phải là object.",
                code="row_not_object",
                row_index=index,
            )
        keys = set(item)
        if keys != v2_keys:
            missing = sorted(v2_keys - keys)
            extra = sorted(str(key) for key in keys - v2_keys)
            details = []
            if missing:
                details.append("thiếu " + ", ".join(missing))
            if extra:
                details.append("thừa " + ", ".join(extra))
            raise SchemaError(
                f"Dòng {index + 1} của schema v2 không đúng trường"
                + (f" ({'; '.join(details)})" if details else "."),
                code="row_keys",
                row_index=index,
            )
        rows.append(DataRow.from_mapping(item))
    return BatchDocument(v=SCHEMA_VERSION, rows=rows)


def document_to_dict(document: BatchDocument) -> dict[str, Any]:
    """Luôn serialize thành schema v2 dạng object."""

    if type(document.v) is not int or document.v not in {
        LEGACY_SCHEMA_VERSION,
        SCHEMA_VERSION,
    }:
        raise SchemaError(
            "Phiên bản tài liệu phải là số nguyên 1 hoặc 2.",
            code="invalid_version",
            field="v",
        )
    rows: list[dict[str, Any]] = []
    for index, row in enumerate(document.rows):
        if not isinstance(row, DataRow):
            raise SchemaError(
                f"Dòng {index + 1} không phải DataRow.",
                code="invalid_internal_row",
                row_index=index,
            )
        rows.append(row.to_object())
    return {"v": SCHEMA_VERSION, "d": rows}


def coerce_document(
    value: BatchDocument | Mapping[str, Any] | Iterable[DataRow | Sequence[Any]],
) -> BatchDocument:
    """Nhận document, root dict hoặc iterable dòng để hỗ trợ các adapter UI."""

    if isinstance(value, BatchDocument):
        return value
    if isinstance(value, Mapping):
        return parse_document(value)
    if isinstance(value, (str, bytes, bytearray)):
        raise SchemaError("Dữ liệu tài liệu không hợp lệ.", code="invalid_document")
    try:
        iterator = iter(value)
    except TypeError as exc:
        raise SchemaError(
            "Dữ liệu tài liệu không phải một iterable dòng hợp lệ.",
            code="invalid_document",
        ) from exc
    rows: list[DataRow] = []
    for index, item in enumerate(iterator):
        try:
            if isinstance(item, DataRow):
                row = item
            elif isinstance(item, Mapping):
                row = DataRow.from_mapping(item)
            else:
                row = DataRow.from_sequence(item)
        except (TypeError, ValueError) as exc:
            raise SchemaError(
                f"Dòng {index + 1} phải là một mảng có đúng 7 phần tử.",
                code="row_length",
                row_index=index,
            ) from exc
        rows.append(row)
    return BatchDocument(rows=rows)


def is_non_boolean_integer(value: object) -> bool:
    return type(value) is int


# Alias ngắn thường được dùng trong test/tích hợp.
parse_schema = parse_document
to_json_object = document_to_dict
