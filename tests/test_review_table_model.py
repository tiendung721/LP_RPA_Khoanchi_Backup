from __future__ import annotations

from PySide6.QtCore import Qt

from app.ui.review_table_model import (
    ReviewFilterProxyModel,
    ReviewRow,
    ReviewTableModel,
    RowStatus,
)
from app.ui.edit_row_dialog import normalize_bl, normalize_container, parse_amount


def _rows() -> list[list[object]]:
    return [
        ["DRYU3026167", None, "VTN", "CV", None, None, 13_554_000],
        [None, "BL123456789", "CB", "HD", None, None, 27_500_000],
        ["ABCD1234567", "HBL-01", "VSDL", "ST", None, None, 850_000],
    ]


def test_sort_filter_does_not_change_source_order(qtbot) -> None:
    model = ReviewTableModel(_rows())
    proxy = ReviewFilterProxyModel()
    proxy.setSourceModel(model)

    original = model.rows_as_arrays()
    proxy.sort(ReviewTableModel.COLUMN_AMOUNT, Qt.SortOrder.DescendingOrder)
    proxy.set_search_text("BL123")

    assert proxy.rowCount() == 1
    assert model.rows_as_arrays() == original


def test_document_filter_combines_with_search_and_keeps_duplicate_names_separate(
    qtbot,
) -> None:
    model = ReviewTableModel(
        [
            ReviewRow(
                cont="DRYU3026167",
                bl="BL-A",
                fee="VTN",
                rule="ST",
                amount=100,
                source_document_id="DOC_001",
                source_document_name="Hoa_don.pdf",
            ),
            ReviewRow(
                cont="MSCU1234567",
                bl="BL-B",
                fee="HH",
                rule="ST",
                amount=200,
                source_document_id="DOC_002",
                source_document_name="Hoa_don.pdf",
            ),
        ]
    )
    proxy = ReviewFilterProxyModel()
    proxy.setSourceModel(model)

    proxy.set_document_filter("DOC_002")
    assert proxy.rowCount() == 1
    assert proxy.data(proxy.index(0, ReviewTableModel.COLUMN_CONT)) == "MSCU1234567"

    proxy.set_search_text("BL-A")
    assert proxy.rowCount() == 0
    proxy.set_search_text("BL-B")
    assert proxy.rowCount() == 1


def test_v1_fields_are_serialized_and_searchable(qtbot) -> None:
    model = ReviewTableModel(
        [
            [
                "DRYU3026167",
                None,
                "VTN",
                "CV",
                "HD-000130",
                "Vận tải Ánh Dương",
                13_554_000,
            ]
        ]
    )
    proxy = ReviewFilterProxyModel()
    proxy.setSourceModel(model)

    serialized = model.to_document()
    assert serialized["v"] == 3
    assert serialized["d"][0]["container"] == "DRYU3026167"
    assert serialized["d"][0]["invoice_no"] == "HD-000130"
    assert serialized["d"][0]["container_count_basis"] == "UNKNOWN"
    assert model.data(model.index(0, ReviewTableModel.COLUMN_INVOICE_NO)) == "HD-000130"
    assert model.data(model.index(0, ReviewTableModel.COLUMN_CARRIER)) == "Vận tải Ánh Dương"

    proxy.set_search_text("ánh dương")
    assert proxy.rowCount() == 1


def test_edit_updates_validation_and_dirty_state(qtbot) -> None:
    model = ReviewTableModel(_rows())
    assert not model.dirty

    model.update_row(0, ReviewRow("DRYU3026167", None, "CB", "CV", 10))

    assert model.dirty
    assert model.validation_at(0).status is RowStatus.ERROR
    assert model.first_error_row() == 0


def test_remove_rows_handles_non_contiguous_source_positions_once(qtbot) -> None:
    model = ReviewTableModel(_rows())
    rows_changed_count = 0

    def count_change() -> None:
        nonlocal rows_changed_count
        rows_changed_count += 1

    model.rowsChanged.connect(count_change)

    removed = model.remove_rows([2, 0, 2])

    assert [row.cont for row in removed] == ["DRYU3026167", "ABCD1234567"]
    assert model.rowCount() == 1
    assert model.row_at(0).bl == "BL123456789"
    assert model.dirty
    assert rows_changed_count == 1


def test_duplicate_rows_are_only_a_warning(qtbot) -> None:
    row = ["DRYU3026167", None, "VTN", "CV", None, None, 13_554_000]
    model = ReviewTableModel([row, row])

    assert model.stats.error == 0
    assert model.stats.warning == 2
    assert model.rowCount() == 2


def test_friendly_amount_parser_and_text_normalization() -> None:
    for text in ("13554000", "13.554.000", "13,554,000", "13 554 000"):
        assert parse_amount(text) == 13_554_000

    assert normalize_container(" dryu-302 6167 ") == "DRYU3026167"
    assert normalize_container(" oolu-0o1i8b7 ") == "OOLU0O1I8B7"
    assert normalize_bl("  hbl /  2026-01  ") == "HBL / 2026-01"


def test_bang_ke_review_allows_negative_and_uses_signed_total(qtbot) -> None:
    row = ReviewRow("DRYU3045911", "VS26020793", "VSDL", "GV", -282_000)

    regular = ReviewTableModel([row])
    bang_ke = ReviewTableModel([row], allow_negative=True)

    assert regular.stats.error == 1
    assert bang_ke.stats.error == 0
    assert bang_ke.stats.warning == 1
    assert bang_ke.stats.total_amount == -282_000
    assert any(
        "Khoản điều chỉnh giảm" in warning
        for warning in bang_ke.validation_at(0).warnings
    )
    assert parse_amount("-282.000", allow_negative=True) == -282_000


def test_invalid_unhashable_values_remain_visible_and_editable(qtbot) -> None:
    model = ReviewTableModel(
        [
            [["OCR"], None, "VTN", "CV", None, None, {"raw": "13.554.000"}],
            [["OCR"], None, "VTN", "CV", None, None, {"raw": "13.554.000"}],
        ]
    )

    assert model.stats.error == 2
    assert model.data(model.index(0, ReviewTableModel.COLUMN_AMOUNT)) == (
        "{'raw': '13.554.000'}"
    )
    assert model.rowCount() == 2


def test_compact_vessel_column_and_action_visibility(qtbot) -> None:
    sea_missing = ReviewRow(
        cont=None,
        bl="BL-1",
        fee="CB",
        rule="HD",
        amount=100,
        vessel_name="NEW VISION",
        voyage_no="2610S",
    )
    sea_with_container = ReviewRow(
        cont="DRYU3026167",
        fee="CB",
        rule="HD",
        amount=100,
    )
    model = ReviewTableModel([sea_missing, sea_with_container])

    assert model.data(model.index(0, ReviewTableModel.COLUMN_VESSEL_VOYAGE)) == (
        "NEW VISION 2610S"
    )
    assert model.data(
        model.index(0, ReviewTableModel.COLUMN_LOOKUP_ACTION),
        ReviewTableModel.ACTION_VISIBLE_ROLE,
    )
    assert not model.data(
        model.index(1, ReviewTableModel.COLUMN_LOOKUP_ACTION),
        ReviewTableModel.ACTION_VISIBLE_ROLE,
    )
