from __future__ import annotations

import json

from app.models import DataRow
from app.schema import parse_document
from app.services.json_codec import JsonCodec
from app.ui.edit_row_dialog import EditRowDialog
from app.ui.review_merge import rows_match
from app.ui.review_table_model import ReviewFilterProxyModel, ReviewRow, ReviewTableModel


def test_v4_source_fee_name_round_trip_and_old_rows_still_load() -> None:
    old = DataRow("DRYU3026167", None, "CXD", "GV", 200_000).to_object()
    named = {**old, "source_fee_name": "PHÍ CHIẾM DỤNG VỎ"}
    document = JsonCodec().loads(
        json.dumps({"v": 4, "d": [old, named]}, ensure_ascii=False)
    )

    assert document.rows[0].source_fee_name is None
    assert document.rows[1].source_fee_name == "PHÍ CHIẾM DỤNG VỎ"
    reloaded = parse_document(json.loads(JsonCodec().dumps(document)))
    assert [row.source_fee_name for row in reloaded.rows] == [
        None, "PHÍ CHIẾM DỤNG VỎ"
    ]


def test_review_shows_source_names_and_preserves_them_when_fee_changes(qtbot) -> None:
    rows = [
        ReviewRow(
            cont="DRYU3026167", fee="CXD", rule="GV", amount=199_800,
            source_fee_name="PHỤ THU CƯỚC",
        ),
        ReviewRow(
            cont="DRYU3026167", fee="CXD", rule="GV", amount=200_000,
            source_fee_name="PHÍ CHIẾM DỤNG VỎ",
        ),
    ]
    model = ReviewTableModel(rows)
    assert model.data(model.index(0, model.COLUMN_SOURCE_FEE_NAME)) == "PHỤ THU CƯỚC"
    assert model.data(model.index(1, model.COLUMN_SOURCE_FEE_NAME)) == "PHÍ CHIẾM DỤNG VỎ"

    proxy = ReviewFilterProxyModel()
    proxy.setSourceModel(model)
    proxy.set_search_text("chiếm dụng")
    assert proxy.rowCount() == 1

    dialog = EditRowDialog(model.row_at(0))
    qtbot.addWidget(dialog)
    assert dialog.source_fee_name_label.text() == "PHỤ THU CƯỚC"
    dialog.fee_combo.setCurrentIndex(dialog.fee_combo.findData("LC"))
    model.update_row(0, dialog.row_data())
    assert model.row_at(0).fee == "LC"
    assert model.row_at(0).source_fee_name == "PHỤ THU CƯỚC"
    assert model.to_document()["d"][0]["source_fee_name"] == "PHỤ THU CƯỚC"


def test_different_source_fee_names_do_not_merge_after_reclassification() -> None:
    base = dict(
        cont="DRYU3026167", fee="LC", rule="GV", amount=200_000,
        invoice_no="HD-1",
    )
    left = ReviewRow(**base, source_fee_name="PHỤ THU CƯỚC")
    right = ReviewRow(**base, source_fee_name="PHÍ CHIẾM DỤNG VỎ")
    assert not rows_match(left, right)
