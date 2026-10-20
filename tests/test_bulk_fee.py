from __future__ import annotations

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog

from app.services.json_codec import JsonCodec
from app.ui.bulk_fee_dialog import BulkFeeDialog
from app.ui.review_table_model import (
    ReviewRow,
    ReviewTableModel,
    bulk_fee_change_positions,
)
from app.ui.review_window import ReviewWindow


def _rows() -> list[ReviewRow]:
    return [
        ReviewRow(
            cont="ABCD1234567", fee="CXD", rule="GV", amount=100,
            source_document_id="DOC_A", source_document_name="a.pdf",
        ),
        ReviewRow(
            cont="EFGH1234567", fee="HV", rule="ST", amount=200,
            carrier="Vận tải cũ", source_document_id="DOC_A",
            source_document_name="a.pdf",
        ),
        ReviewRow(
            cont="IJKL1234567", fee="CB", rule="HD", amount=300,
            carrier="Tàu biển", source_document_id="DOC_A",
            source_document_name="a.pdf",
        ),
        ReviewRow(
            cont="MNOP1234567", fee="NV", rule="HD", amount=400,
            carrier="Vận tải khác", source_document_id="DOC_B",
            source_document_name="b.pdf",
        ),
    ]


def _window(qtbot) -> ReviewWindow:
    window = ReviewWindow(
        {
            "metadata": {"id": 12, "status": "REVIEWING", "source_kind": "ASSISTANT"},
            "document": {"v": 4, "d": [row.to_object() for row in _rows()]},
        }
    )
    qtbot.addWidget(window)
    return window


def test_bulk_fee_planning_excludes_cb_and_validates_modes() -> None:
    rows = _rows()
    assert bulk_fee_change_positions(rows, range(4), fee="HH", mode="fill") == [0]
    assert bulk_fee_change_positions(
        rows, range(4), fee="HH", mode="replace", old_fee="HV"
    ) == [1]
    assert bulk_fee_change_positions(
        rows, range(4), fee="HH", mode="overwrite"
    ) == [0, 1, 3]
    with pytest.raises(ValueError):
        bulk_fee_change_positions(rows, range(4), fee="CB", mode="overwrite")
    with pytest.raises(ValueError):
        bulk_fee_change_positions(
            rows, range(4), fee="HH", mode="replace", old_fee="CB"
        )


def test_bulk_fee_dialog_scopes_and_preview(qtbot) -> None:
    dialog = BulkFeeDialog(
        _rows(), selected_positions=[1, 2], filtered_positions=[0, 1, 2]
    )
    qtbot.addWidget(dialog)
    assert dialog.scope_combo.currentData() == "selected"
    assert dialog.new_fee_combo.findData("CB") == -1
    assert not dialog.apply_button.isEnabled()

    dialog.new_fee_combo.setCurrentIndex(dialog.new_fee_combo.findData("HH"))
    dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("overwrite"))
    assert dialog.change_positions() == [1]
    assert "1 dòng CB" in dialog.preview_label.text()
    assert "1 dòng" in dialog.preview_label.text()
    assert dialog.apply_button.isEnabled()

    dialog.scope_combo.setCurrentIndex(dialog.scope_combo.findData("filtered"))
    assert dialog.change_positions() == [0, 1]
    assert 3 not in dialog.scope_positions()


def test_bulk_fee_model_clears_carrier_and_invalid_hd_rule(qtbot) -> None:
    model = ReviewTableModel(_rows())
    runtime_ids = [row.runtime_id for row in model.rows()]
    assert model.bulk_update_fee(range(4), fee="HH", mode="overwrite") == 3
    rows = model.rows()
    assert [row.fee for row in rows] == ["HH", "HH", "CB", "HH"]
    assert [row.carrier for row in rows] == [None, None, "Tàu biển", None]
    assert [row.rule for row in rows] == ["GV", "ST", "HD", None]
    assert [row.runtime_id for row in rows] == runtime_ids
    assert model.dirty
    model.mark_clean()
    assert model.bulk_update_fee(range(4), fee="HH", mode="overwrite") == 0
    assert not model.dirty


def test_bulk_fee_window_uses_selected_source_row_after_sort(qtbot, monkeypatch) -> None:
    window = _window(qtbot)
    window.proxy_model.sort(ReviewTableModel.COLUMN_AMOUNT, Qt.SortOrder.DescendingOrder)
    source_row = 1
    proxy_index = window.proxy_model.mapFromSource(
        window.model.index(source_row, ReviewTableModel.COLUMN_NO)
    )
    window.table.selectRow(proxy_index.row())

    def accept_selected(dialog: BulkFeeDialog) -> QDialog.DialogCode:
        dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("overwrite"))
        dialog.new_fee_combo.setCurrentIndex(dialog.new_fee_combo.findData("HH"))
        assert dialog.scope_combo.currentData() == "selected"
        assert dialog.change_positions() == [source_row]
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(BulkFeeDialog, "exec", accept_selected)
    window.bulk_edit_fees()
    assert [row.fee for row in window.model.rows()] == ["CXD", "HH", "CB", "NV"]
    assert window.model.row_at(1).carrier is None
    window.model.mark_clean()
    window.close()


def test_bulk_fee_filtered_scope_excludes_hidden_rows(qtbot, monkeypatch, tmp_path) -> None:
    window = _window(qtbot)
    window.document_filter.setCurrentIndex(window.document_filter.findData("DOC_A"))
    assert window.proxy_model.rowCount() == 3

    def accept_filtered(dialog: BulkFeeDialog) -> QDialog.DialogCode:
        dialog.scope_combo.setCurrentIndex(dialog.scope_combo.findData("filtered"))
        dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("overwrite"))
        dialog.new_fee_combo.setCurrentIndex(dialog.new_fee_combo.findData("HH"))
        assert dialog.change_positions() == [0, 1]
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(BulkFeeDialog, "exec", accept_filtered)
    window.bulk_edit_fees()
    assert [row.fee for row in window.model.rows()] == ["HH", "HH", "CB", "NV"]

    output_path = tmp_path / "reviewed.json"
    codec = JsonCodec()
    window.set_save_handler(
        lambda _batch_id, document: codec.dump_atomic(output_path, document)
    )
    assert window.save_working()
    assert [row.fee for row in codec.load(output_path).rows] == ["HH", "HH", "CB", "NV"]
    window.close()
