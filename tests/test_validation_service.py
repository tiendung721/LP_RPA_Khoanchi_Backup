from __future__ import annotations

import pytest

from app.models import BatchDocument, DataRow, Severity
from app.services.validation_service import (
    AmountParseError,
    ValidationService,
    normalize_bl,
    normalize_bl_key,
    normalize_container,
    normalize_optional_text,
    parse_amount,
)


@pytest.fixture
def validator() -> ValidationService:
    return ValidationService()


@pytest.mark.parametrize(
    ("row", "code"),
    [
        (DataRow("DRYU3026167", None, "LA", "CV", 1), "fee_unknown"),
        (DataRow("DRYU3026167", None, "VTN", "XX", 1), "rule_unknown"),
        (DataRow("DRYU3026167", None, "VTN", "CV", "1"), "amount_type"),
        (DataRow("DRYU3026167", None, "VTN", "CV", True), "amount_boolean"),
        (DataRow("DRYU3026167", None, "VTN", "CV", -1), "amount_negative"),
        (DataRow("DRYU3026167", None, "VTN", "HD", 1), "hd_requires_cb"),
        (DataRow("DRYU3026167", "BL1", "CB", "CV", 1), "cb_requires_hd"),
        (
            DataRow(
                "DRYU3026167",
                None,
                "VTN",
                "CV",
                1,
                invoice_no=130,  # type: ignore[arg-type]
            ),
            "invoice_no_type",
        ),
        (
            DataRow(
                "DRYU3026167",
                None,
                "VTN",
                "CV",
                1,
                carrier=123,  # type: ignore[arg-type]
            ),
            "carrier_type",
        ),
    ],
)
def test_blocking_row_rules(
    validator: ValidationService, row: DataRow, code: str
) -> None:
    result = validator.validate_row(row)

    assert result.status is Severity.ERROR
    assert code in {issue.code for issue in result.issues}


def test_container_format_is_only_warning(validator: ValidationService) -> None:
    valid = validator.validate_row(DataRow("DRYU3026167", None, "VTN", "CV", 1))
    unusual = validator.validate_row(DataRow("DRYUO026167", None, "VTN", "CV", 1))

    assert "cont_format" not in {issue.code for issue in valid.issues}
    assert unusual.status is Severity.WARNING
    assert "cont_format" in {issue.code for issue in unusual.issues}


def test_normalization_never_guesses_ocr_characters() -> None:
    assert normalize_container(" oolu-0o1 i8b7 ") == "OOLU0O1I8B7"
    assert normalize_container("   ") is None
    assert normalize_bl("  hbl /  2026-01  ") == "HBL / 2026-01"
    assert normalize_bl_key("  hbl /  2026-01  ") == "HBL202601"
    assert normalize_bl_key("vs 2607-1443") == "VS26071443"
    assert normalize_bl("") is None
    assert normalize_optional_text("  000130 / HD  ") == "000130 / HD"
    assert normalize_optional_text("   ") is None


@pytest.mark.parametrize(
    "text", ["13554000", "13.554.000", "13,554,000", "13 554 000"]
)
def test_amount_parser_accepts_friendly_thousands(text: str) -> None:
    assert parse_amount(text) == 13_554_000


@pytest.mark.parametrize(
    "value", ["13.5", "1,23", "-1", "1 000.000", "₫1000", True, -1]
)
def test_amount_parser_rejects_ambiguous_or_invalid_values(value: object) -> None:
    with pytest.raises(AmountParseError):
        parse_amount(value)  # type: ignore[arg-type]


def test_bang_ke_validation_keeps_signed_adjustment() -> None:
    row = DataRow("DRYU3045911", "VS26020793", "VSDL", "GV", -282_000)
    validator = ValidationService()

    regular = validator.validate_document(BatchDocument(rows=[row]))
    bang_ke = validator.validate_document(
        BatchDocument(rows=[row]), allow_negative=True
    )

    assert regular.has_errors
    assert not bang_ke.has_errors
    assert bang_ke.summary.total_amount == -282_000
    assert any(
        issue.code == "amount_negative_adjustment"
        for issue in bang_ke.row_results[0].issues
    )


def test_duplicate_rows_remain_and_only_warn(
    validator: ValidationService,
) -> None:
    row = DataRow("DRYU3026167", None, "VTN", "CV", 1)
    document = BatchDocument(rows=[row, row.copy_with()])

    result = validator.validate_document(document)

    assert len(document.rows) == 2
    assert result.summary.error_count == 0
    assert result.summary.warning_count == 2
    assert all(
        "duplicate_row" in {issue.code for issue in row_result.issues}
        for row_result in result.row_results
    )


def test_file_summary_counts_and_total(validator: ValidationService) -> None:
    document = BatchDocument(
        rows=[
            DataRow("DRYU3026167", None, "VTN", "CV", 13_554_000),
            DataRow(None, "BL123", "CB", "HD", 27_500_000),
            DataRow(None, None, "CXD", None, None),
        ]
    )

    summary = validator.validate_document(document).summary

    assert summary.total_rows == 3
    assert summary.container_count == 1
    assert summary.bl_count == 1
    assert summary.amount_count == 2
    assert summary.total_amount == 41_054_000
    assert summary.fee_counts == {"CB": 1, "CXD": 1, "VTN": 1}
