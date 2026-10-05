from __future__ import annotations

from app.ui.review_merge import (
    candidate_groups,
    conflicting_fields,
    has_valid_amounts,
    is_exact_duplicate_group,
    merged_row,
    rows_match,
)
from app.ui.review_table_model import ReviewRow


def row(
    amount: int | None,
    *,
    invoice_no: str | None = "00026955",
    invoice_date: str | None = "2026-08-29",
    fee: str = "VSDL",
    cont: str | None = "DRYU3026167",
    **changes: object,
) -> ReviewRow:
    return ReviewRow(
        amount=amount,
        invoice_no=invoice_no,
        invoice_date=invoice_date,
        fee=fee,
        cont=cont,
        rule="ST",
        source_document_id="DOC_001",
        source_document_name="Hoa_don.pdf",
        **changes,
    )


def test_invoice_or_date_match_and_empty_values_do_not_match() -> None:
    base = row(100)
    assert rows_match(base, row(200, invoice_no=" 00026955 ", invoice_date=None))
    assert rows_match(base, row(200, invoice_no="DIFFERENT", invoice_date="2026-08-29"))
    assert not rows_match(base, row(200, invoice_no="26955", invoice_date=None))
    assert not rows_match(base, row(200, invoice_no=None, invoice_date=None))
    assert not rows_match(base, row(200, fee="HV"))
    assert not rows_match(base, row(200, cont=None))
    assert not rows_match(base, row(200, fee="CXD"))


def test_groups_do_not_chain_through_a_middle_row() -> None:
    rows = [
        row(100, invoice_no="A", invoice_date="2026-08-28"),
        row(200, invoice_no="A", invoice_date="2026-08-29"),
        row(300, invoice_no="B", invoice_date="2026-08-29"),
    ]
    assert candidate_groups(rows) == [(0, 1)]


def test_merged_line_retains_original_match_boundaries() -> None:
    first = row(100, invoice_no="A", invoice_date="2026-08-28")
    second = row(200, invoice_no="A", invoice_date="2026-08-29")
    third = row(300, invoice_no="B", invoice_date="2026-08-29")
    combined = merged_row([first, second], {"invoice_date": "2026-08-29"})
    assert rows_match(combined, third)
    assert candidate_groups(
        [combined, third],
        origins={combined.runtime_id: (first, second)},
    ) == []


def test_merge_keeps_fee_rule_and_first_runtime_id_and_fills_known_fields() -> None:
    first = row(291_600, invoice_date=None, carrier=None)
    second = row(70_200, carrier="VT A")
    third = row(75_600, carrier="VT A")
    group = [first, second, third]
    assert conflicting_fields(group) == ()
    result = merged_row(group)
    assert result.amount == 437_400
    assert result.fee == "VSDL"
    assert result.rule == "ST"
    assert result.invoice_date == "2026-08-29"
    assert result.carrier == "VT A"
    assert result.runtime_id == first.runtime_id


def test_conflicts_and_exact_duplicates_need_review() -> None:
    left = row(100, invoice_no="INV-A", sqt=1)
    right = row(200, invoice_no="INV-B", sqt=2)
    assert rows_match(left, right)  # Cùng ngày HĐ.
    assert conflicting_fields([left, right]) == ("invoice_no", "sqt")
    assert merged_row([left, right], {"invoice_no": "INV-B", "sqt": 2}).amount == 300
    assert is_exact_duplicate_group([left, row(100, invoice_no="INV-A", sqt=1)])
    assert not has_valid_amounts([left, row(None)])


def test_rule_conflict_keeps_the_user_selected_rule_and_cb_uses_hd() -> None:
    first = row(100)
    second = row(200)
    second.rule = "GV"
    assert "rule" in conflicting_fields([first, second])
    assert merged_row([first, second], {"rule": "GV"}).rule == "GV"

    cb_first = row(100, fee="CB")
    cb_second = row(200, fee="CB")
    cb_first.rule = "HD"
    cb_second.rule = "HD"
    assert merged_row([cb_first, cb_second]).rule == "HD"


def test_invalid_fee_does_not_crash_matching() -> None:
    assert not rows_match(row(100), row(200, fee=["VSDL"]))
