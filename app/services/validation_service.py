"""Chuẩn hóa dữ liệu người dùng nhập và kiểm tra toàn bộ lô JSON."""

from __future__ import annotations

import json
import re
import unicodedata
from collections import Counter
from collections.abc import Mapping, Sequence
from datetime import date
from typing import Any

from app.constants import (
    CONTAINER_PATTERN,
    FEE_CODES,
    RULE_CODES,
    SCHEMA_VERSION,
)
from app.models import (
    BatchDocument,
    DataRow,
    RowValidation,
    Severity,
    ValidationIssue,
    ValidationResult,
    ValidationSummary,
)
from app.schema import SchemaError, parse_document

_CONTAINER_RE = re.compile(CONTAINER_PATTERN)
_CONTAINER_SEPARATORS_RE = re.compile(r"[\s\-_–—]+", flags=re.UNICODE)
_COLLAPSE_WHITESPACE_RE = re.compile(r"\s+", flags=re.UNICODE)
_AMOUNT_ALLOWED_RE = re.compile(r"^[0-9][0-9.,\s]*$", flags=re.UNICODE)
_AMOUNT_GROUPED_RE = re.compile(r"^[0-9]{1,3}(?:([.,\s])[0-9]{3})(?:\1[0-9]{3})*$")
_MAX_SIGNED_64 = 9_223_372_036_854_775_807
_CONTAINER_COUNT_BASES = frozenset({"EXPLICIT", "CALCULATED", "UNKNOWN"})


class AmountParseError(ValueError):
    """Số tiền người dùng nhập không thể chuẩn hóa an toàn."""


def normalize_container(value: str | None) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise TypeError("Container phải là chuỗi hoặc để trống.")
    normalized = _CONTAINER_SEPARATORS_RE.sub("", value.strip().upper())
    return normalized or None


def normalize_bl(value: str | None) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise TypeError("B/L phải là chuỗi hoặc để trống.")
    normalized = _COLLAPSE_WHITESPACE_RE.sub(" ", value.strip()).upper()
    return normalized or None


def normalize_bl_key(value: str | None) -> str | None:
    """Tạo khóa B/L ổn định để đối chiếu mà không đổi giá trị hiển thị."""

    normalized = normalize_bl(value)
    if normalized is None:
        return None
    normalized = unicodedata.normalize("NFKC", normalized)
    key = re.sub(r"[^A-Z0-9]", "", normalized.upper())
    return key or None


def normalize_fee(value: str) -> str:
    if not isinstance(value, str):
        raise TypeError("Mã loại cước phải là chuỗi.")
    normalized = value.strip().upper()
    if normalized not in FEE_CODES:
        raise ValueError("Mã loại cước không thuộc danh mục chính thức.")
    return normalized


def normalize_rule(value: str | None) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise TypeError("Mã xử lý tiền phải là chuỗi hoặc để trống.")
    normalized = value.strip().upper()
    if not normalized:
        return None
    if normalized not in RULE_CODES:
        raise ValueError("Mã xử lý tiền không thuộc danh mục chính thức.")
    return normalized


def normalize_optional_text(
    value: str | None,
    *,
    field_name: str = "Giá trị",
) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise TypeError(f"{field_name} phải là chuỗi hoặc để trống.")
    normalized = _COLLAPSE_WHITESPACE_RE.sub(" ", value.strip())
    return normalized or None


def parse_amount(
    value: int | str | None,
    *,
    allow_negative: bool = False,
) -> int | None:
    """Parse số nguyên VND; không suy đoán ký hiệu thập phân."""

    if value is None:
        return None
    if isinstance(value, bool):
        raise AmountParseError("Số tiền không được là kiểu boolean.")
    if type(value) is int:
        if value < 0 and not allow_negative:
            raise AmountParseError("Số tiền không được âm.")
        if abs(value) > _MAX_SIGNED_64:
            raise AmountParseError("Số tiền vượt giới hạn số nguyên 64 bit.")
        return value
    if not isinstance(value, str):
        raise AmountParseError("Số tiền phải là số nguyên hoặc để trống.")

    text = value.strip()
    if not text:
        return None
    sign = -1 if text.startswith("-") else 1
    unsigned = text[1:].strip() if sign < 0 else text
    if sign < 0 and not allow_negative:
        raise AmountParseError("Số tiền không được âm.")
    if not _AMOUNT_ALLOWED_RE.fullmatch(unsigned):
        raise AmountParseError(
            "Số tiền chỉ được chứa chữ số và dấu phân cách hàng nghìn."
        )
    if unsigned.isdigit():
        parsed = int(unsigned)
    else:
        # Thu gọn mọi dạng khoảng trắng thành dấu cách trước khi kiểm tra nhóm.
        grouped = _COLLAPSE_WHITESPACE_RE.sub(" ", unsigned)
        match = _AMOUNT_GROUPED_RE.fullmatch(grouped)
        if match is None:
            raise AmountParseError(
                "Dấu phân cách tiền không đúng nhóm hàng nghìn."
            )
        parsed = int(re.sub(r"[.,\s]", "", grouped))
    parsed *= sign
    if abs(parsed) > _MAX_SIGNED_64:
        raise AmountParseError("Số tiền vượt giới hạn số nguyên 64 bit.")
    return parsed


class ValidationService:
    """Kiểm tra schema, kiểu trường, quan hệ nghiệp vụ và cảnh báo."""

    def normalize_row(
        self,
        row: DataRow | None = None,
        *,
        cont: str | None = None,
        bl: str | None = None,
        fee: str | None = None,
        rule: str | None = None,
        invoice_no: str | None = None,
        carrier: str | None = None,
        vessel_voyage_raw: str | None = None,
        vessel_name: str | None = None,
        voyage_no: str | None = None,
        invoice_container_count: int | None = None,
        container_count_basis: str = "UNKNOWN",
        invoice_date: str | None = None,
        source_document_id: str = "MANUAL",
        source_document_name: str = "Dòng thêm thủ công",
        amount: int | str | None = None,
    ) -> DataRow:
        if row is not None:
            cont = row.cont
            bl = row.bl
            fee = row.fee
            rule = row.rule
            invoice_no = row.invoice_no
            carrier = row.carrier
            vessel_voyage_raw = row.vessel_voyage_raw
            vessel_name = row.vessel_name
            voyage_no = row.voyage_no
            invoice_container_count = row.invoice_container_count
            container_count_basis = row.container_count_basis
            invoice_date = row.invoice_date
            source_document_id = row.source_document_id
            source_document_name = row.source_document_name
            amount = row.amount
        if fee is None:
            raise ValueError("Vui lòng chọn mã loại cước.")
        return DataRow(
            cont=normalize_container(cont),
            bl=normalize_bl(bl),
            fee=normalize_fee(fee),
            rule=normalize_rule(rule),
            amount=parse_amount(amount),
            invoice_no=normalize_optional_text(invoice_no, field_name="Số HĐ"),
            carrier=normalize_optional_text(carrier, field_name="Bên vận tải"),
            vessel_voyage_raw=normalize_optional_text(
                vessel_voyage_raw, field_name="Tàu/chuyến nguyên văn"
            ),
            vessel_name=normalize_optional_text(vessel_name, field_name="Tên tàu"),
            voyage_no=normalize_optional_text(voyage_no, field_name="Số chuyến"),
            invoice_container_count=invoice_container_count,
            container_count_basis=(container_count_basis or "UNKNOWN").strip().upper(),
            invoice_date=normalize_optional_text(invoice_date, field_name="Ngày HĐ"),
            source_document_id=normalize_optional_text(
                source_document_id, field_name="Mã chứng từ nguồn"
            ) or "",
            source_document_name=normalize_optional_text(
                source_document_name, field_name="Tên chứng từ nguồn"
            ) or "",
        )

    def validate_row(
        self,
        row: DataRow,
        index: int = 0,
        *,
        allow_negative: bool = False,
    ) -> RowValidation:
        issues: list[ValidationIssue] = []

        def error(code: str, message: str, field: str) -> None:
            issues.append(
                ValidationIssue(Severity.ERROR, code, message, index, field)
            )

        def warning(code: str, message: str, field: str | None = None) -> None:
            issues.append(
                ValidationIssue(Severity.WARNING, code, message, index, field)
            )

        cont_valid_type = row.cont is None or isinstance(row.cont, str)
        bl_valid_type = row.bl is None or isinstance(row.bl, str)
        fee_valid = isinstance(row.fee, str) and row.fee in FEE_CODES
        rule_valid = row.rule is None or (
            isinstance(row.rule, str) and row.rule in RULE_CODES
        )
        invoice_no_valid_type = row.invoice_no is None or isinstance(
            row.invoice_no, str
        )
        carrier_valid_type = row.carrier is None or isinstance(row.carrier, str)
        amount_valid_type = row.amount is None or type(row.amount) is int
        raw_valid_type = row.vessel_voyage_raw is None or isinstance(
            row.vessel_voyage_raw, str
        )
        vessel_valid_type = row.vessel_name is None or isinstance(
            row.vessel_name, str
        )
        voyage_valid_type = row.voyage_no is None or isinstance(row.voyage_no, str)
        count_valid_type = (
            row.invoice_container_count is None
            or type(row.invoice_container_count) is int
        )
        basis_valid = (
            isinstance(row.container_count_basis, str)
            and row.container_count_basis in _CONTAINER_COUNT_BASES
        )
        invoice_date_valid_type = row.invoice_date is None or isinstance(
            row.invoice_date, str
        )
        source_id_valid = (
            isinstance(row.source_document_id, str)
            and bool(row.source_document_id.strip())
        )
        source_name_valid = (
            isinstance(row.source_document_name, str)
            and bool(row.source_document_name.strip())
        )

        if not source_id_valid:
            error(
                "source_document_id_required",
                "Mã chứng từ nguồn phải là chuỗi không rỗng.",
                "source_document_id",
            )
        if not source_name_valid:
            error(
                "source_document_name_required",
                "Tên chứng từ nguồn phải là chuỗi không rỗng.",
                "source_document_name",
            )

        if not cont_valid_type:
            error(
                "cont_type",
                "Container phải là chuỗi hoặc null.",
                "cont",
            )
        elif isinstance(row.cont, str) and not _CONTAINER_RE.fullmatch(row.cont):
            warning(
                "cont_format",
                "Container chưa đúng mẫu 4 chữ cái và 7 chữ số.",
                "cont",
            )

        if not bl_valid_type:
            error("bl_type", "B/L phải là chuỗi hoặc null.", "bl")

        if not fee_valid:
            error(
                "fee_unknown",
                "Mã loại cước không thuộc danh mục chính thức.",
                "fee",
            )

        if not rule_valid:
            error(
                "rule_unknown",
                "Mã xử lý tiền phải là HD, ST, CV, GV hoặc null.",
                "rule",
            )

        if not invoice_no_valid_type:
            error("invoice_no_type", "Số HĐ phải là chuỗi hoặc null.", "invoice_no")

        if not carrier_valid_type:
            error("carrier_type", "Bên vận tải phải là chuỗi hoặc null.", "carrier")

        if not raw_valid_type:
            error(
                "vessel_voyage_raw_type",
                "Tàu/chuyến nguyên văn phải là chuỗi hoặc null.",
                "vessel_voyage_raw",
            )
        if not vessel_valid_type:
            error("vessel_name_type", "Tên tàu phải là chuỗi hoặc null.", "vessel_name")
        if not voyage_valid_type:
            error("voyage_no_type", "Số chuyến phải là chuỗi hoặc null.", "voyage_no")
        if not count_valid_type:
            error(
                "invoice_container_count_type",
                "SL cont HĐ phải là số nguyên hoặc null.",
                "invoice_container_count",
            )
        elif isinstance(row.invoice_container_count, int) and row.invoice_container_count <= 0:
            error(
                "invoice_container_count_positive",
                "SL cont HĐ phải lớn hơn 0.",
                "invoice_container_count",
            )
        if not basis_valid:
            error(
                "container_count_basis_unknown",
                "Căn cứ SL cont phải là EXPLICIT, CALCULATED hoặc UNKNOWN.",
                "container_count_basis",
            )
        elif row.invoice_container_count is None and row.container_count_basis != "UNKNOWN":
            error(
                "container_count_basis_without_count",
                "Không có SL cont thì căn cứ phải là UNKNOWN.",
                "container_count_basis",
            )
        elif row.invoice_container_count is not None and row.container_count_basis == "UNKNOWN":
            error(
                "container_count_missing_basis",
                "Có SL cont thì phải chỉ rõ EXPLICIT hoặc CALCULATED.",
                "container_count_basis",
            )
        if not invoice_date_valid_type:
            error("invoice_date_type", "Ngày HĐ phải là chuỗi hoặc null.", "invoice_date")
        elif isinstance(row.invoice_date, str):
            try:
                parsed_invoice_date = date.fromisoformat(row.invoice_date)
            except ValueError:
                parsed_invoice_date = None
            if parsed_invoice_date is None or parsed_invoice_date.isoformat() != row.invoice_date:
                error(
                    "invoice_date_format",
                    "Ngày HĐ phải đúng định dạng YYYY-MM-DD.",
                    "invoice_date",
                )

        if not amount_valid_type:
            if isinstance(row.amount, bool):
                message = "Số tiền không được là kiểu boolean."
                code = "amount_boolean"
            else:
                message = "Số tiền phải là số nguyên hoặc null."
                code = "amount_type"
            error(code, message, "amount")
        elif isinstance(row.amount, int) and row.amount < 0 and not allow_negative:
            error("amount_negative", "Số tiền không được âm.", "amount")
        elif isinstance(row.amount, int) and row.amount < 0:
            warning(
                "amount_negative_adjustment",
                "Khoản điều chỉnh giảm.",
                "amount",
            )

        if fee_valid and rule_valid:
            if row.rule == "HD" and row.fee != "CB":
                error(
                    "hd_requires_cb",
                    "Quy tắc HD chỉ được dùng cho loại cước CB.",
                    "rule",
                )
            if row.fee == "CB" and row.rule != "HD":
                error(
                    "cb_requires_hd",
                    "Loại cước CB bắt buộc dùng quy tắc HD.",
                    "rule",
                )

        if cont_valid_type and bl_valid_type and row.cont is None and row.bl is None:
            warning(
                "missing_reference",
                "Cả container và B/L đều chưa được xác định.",
            )
        if row.fee == "CXD":
            warning(
                "fee_cxd",
                "Loại cước CXD cần được người dùng kiểm tra lại.",
                "fee",
            )
        if amount_valid_type and row.amount is None:
            warning(
                "amount_missing",
                "Số tiền chưa được xác định.",
                "amount",
            )
        elif amount_valid_type and row.amount == 0:
            warning(
                "amount_zero",
                "Số tiền đang bằng 0.",
                "amount",
            )
        if rule_valid and row.rule is None:
            warning(
                "rule_missing",
                "Quy tắc xử lý tiền chưa được xác định.",
                "rule",
            )
        if (
            fee_valid
            and row.fee == "CB"
            and bl_valid_type
            and row.bl is None
        ):
            warning(
                "cb_missing_bl",
                "Dòng cước biển chưa có B/L.",
                "bl",
            )
        if fee_valid and row.fee == "CB" and row.cont is None:
            required_reconciliation = (
                ("vessel_name", row.vessel_name, "tên tàu"),
                ("voyage_no", row.voyage_no, "số chuyến"),
                (
                    "invoice_container_count",
                    row.invoice_container_count,
                    "SL cont hóa đơn",
                ),
            )
            for field_name, value, label in required_reconciliation:
                if value in (None, ""):
                    warning(
                        f"cb_missing_{field_name}",
                        f"Dòng cước biển thiếu container chưa có {label} để đối soát BK.",
                        field_name,
                    )
        return RowValidation(row_index=index, issues=issues)

    def validate_document(
        self,
        value: BatchDocument | object,
        *,
        allow_negative: bool = False,
    ) -> ValidationResult:
        if isinstance(value, BatchDocument):
            document = value
            if type(document.v) is not int or document.v != SCHEMA_VERSION:
                issue = ValidationIssue(
                    Severity.ERROR,
                    "invalid_version",
                    f"Khóa v nội bộ phải là số nguyên {SCHEMA_VERSION}.",
                    field="v",
                )
                return ValidationResult(
                    issues=[issue],
                    summary=ValidationSummary(total_rows=len(document.rows)),
                )
        else:
            try:
                document = parse_document(value)
            except SchemaError as exc:
                issue = ValidationIssue(
                    Severity.ERROR,
                    exc.code,
                    str(exc),
                    exc.row_index,
                    exc.field,
                )
                return ValidationResult(issues=[issue])

        row_results = [
            self.validate_row(row, index, allow_negative=allow_negative)
            for index, row in enumerate(document.rows)
        ]
        self._append_duplicate_warnings(document.rows, row_results)
        all_issues = [
            issue for row_result in row_results for issue in row_result.issues
        ]

        statuses = [result.status for result in row_results]
        fee_counts: Counter[str] = Counter()
        container_count = 0
        bl_count = 0
        amount_count = 0
        total_amount = 0
        for row in document.rows:
            if isinstance(row.cont, str) and bool(row.cont):
                container_count += 1
            if isinstance(row.bl, str) and bool(row.bl):
                bl_count += 1
            if type(row.amount) is int:
                amount_count += 1
                total_amount += row.amount
            if isinstance(row.fee, str):
                fee_counts[row.fee] += 1

        summary = ValidationSummary(
            total_rows=len(document.rows),
            valid_count=statuses.count(Severity.VALID),
            warning_count=statuses.count(Severity.WARNING),
            error_count=statuses.count(Severity.ERROR),
            container_count=container_count,
            bl_count=bl_count,
            amount_count=amount_count,
            total_amount=total_amount,
            fee_counts=dict(sorted(fee_counts.items())),
        )
        return ValidationResult(
            issues=all_issues,
            row_results=row_results,
            summary=summary,
        )

    @staticmethod
    def _append_duplicate_warnings(
        rows: Sequence[DataRow],
        results: Sequence[RowValidation],
    ) -> None:
        positions: dict[str, list[int]] = {}
        for index, row in enumerate(rows):
            # JSON tạo khóa ổn định cả khi một trường đầu vào sai kiểu/unhashable.
            try:
                key = json.dumps(
                    row.to_list(),
                    ensure_ascii=False,
                    sort_keys=True,
                    separators=(",", ":"),
                )
            except (TypeError, ValueError):
                key = repr(row.to_list())
            positions.setdefault(key, []).append(index)
        for duplicate_indices in positions.values():
            if len(duplicate_indices) < 2:
                continue
            display_rows = ", ".join(str(index + 1) for index in duplicate_indices)
            for index in duplicate_indices:
                results[index].issues.append(
                    ValidationIssue(
                        Severity.WARNING,
                        "duplicate_row",
                        f"Dòng trùng hoàn toàn với các dòng: {display_rows}.",
                        row_index=index,
                    )
                )


def validate_document(value: BatchDocument | Mapping[str, Any]) -> ValidationResult:
    return ValidationService().validate_document(value)


def validate_row(row: DataRow, index: int = 0) -> RowValidation:
    return ValidationService().validate_row(row, index)
