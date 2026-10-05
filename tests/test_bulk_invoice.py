from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog

from app.ui.bulk_invoice_dialog import BulkInvoiceDialog
from app.ui.review_table_model import ReviewRow, ReviewTableModel
from app.ui.review_window import ReviewWindow
from app.services.json_codec import JsonCodec


def _rows() -> list[ReviewRow]:
    return [
        ReviewRow(
            cont="ABCD1234567", fee="VSDL", rule="GV", amount=100,
            source_document_id="BK_A", source_document_name="bang_ke_a.xlsx",
        ),
        ReviewRow(
            cont="EFGH1234567", fee="QT", rule="GV", amount=200,
            invoice_no="HD-CU", source_document_id="BK_A",
            source_document_name="bang_ke_a.xlsx",
        ),
        ReviewRow(
            cont="IJKL1234567", fee="LL", rule="GV", amount=300,
            invoice_no="HD-LL", source_document_id="BK_A",
            source_document_name="bang_ke_a.xlsx",
        ),
        ReviewRow(
            cont="MNOP1234567", fee="GH", rule="GV", amount=400,
            source_document_id="BK_B", source_document_name="bang_ke_b.xlsx",
        ),
    ]


def _window(qtbot) -> ReviewWindow:
    window = ReviewWindow(
        {
            "metadata": {"id": 10, "status": "REVIEWING", "source_kind": "BANG_KE"},
            "document": {"v": 4, "d": [row.to_object() for row in _rows()]},
        }
    )
    qtbot.addWidget(window)
    return window


def test_bulk_invoice_dialog_scopes_preview_and_excludes_ll(qtbot) -> None:
    dialog = BulkInvoiceDialog(_rows(), selected_positions=[1, 2], document_id="BK_A")
    qtbot.addWidget(dialog)
    dialog.new_invoice_edit.setText(" 000123 ")

    assert dialog.scope_combo.currentData() == "selected"
    assert dialog.change_positions() == []
    assert not dialog.apply_button.isEnabled()

    dialog.scope_combo.setCurrentIndex(dialog.scope_combo.findData("all"))
    assert dialog.change_positions() == [0]
    assert "1 dòng" in dialog.preview_label.text()
    assert "HD-CU" in dialog.existing_label.text()
    assert "HD-LL" not in dialog.existing_label.text()
    assert dialog.apply_button.isEnabled()

    dialog.scope_combo.setCurrentIndex(dialog.scope_combo.findData("selected"))
    assert dialog.change_positions() == []
    assert not dialog.apply_button.isEnabled()

    dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("replace"))
    assert dialog.old_invoice_combo.count() == 1
    assert dialog.old_invoice_combo.currentData() == "HD-CU"
    assert dialog.change_positions() == [1]
    assert dialog.apply_button.isEnabled()

    dialog.document_combo.setCurrentIndex(dialog.document_combo.findData("BK_B"))
    dialog.scope_combo.setCurrentIndex(dialog.scope_combo.findData("all"))
    dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("fill"))
    assert dialog.change_positions() == [3]


def test_bulk_invoice_model_preserves_rows_and_skips_noop(qtbot) -> None:
    model = ReviewTableModel(_rows())
    runtime_ids = [model.runtime_id_at(index) for index in range(model.rowCount())]

    assert model.bulk_update_invoice_no(
        [0, 1, 2], invoice_no=" 000123 ", mode="fill"
    ) == 1
    assert [row.invoice_no for row in model.rows()] == [
        "000123", "HD-CU", "HD-LL", None,
    ]
    assert model.dirty
    assert [model.runtime_id_at(index) for index in range(model.rowCount())] == runtime_ids
    assert model.to_document()["d"][0]["invoice_no"] == "000123"

    model.mark_clean()
    assert model.bulk_update_invoice_no(
        [0, 1, 2], invoice_no="000123", mode="fill"
    ) == 0
    assert not model.dirty

    assert model.bulk_update_invoice_no(
        [0, 1, 2], invoice_no="HD-MOI", mode="replace",
        old_invoice_no="hd-cu",
    ) == 1
    assert [row.invoice_no for row in model.rows()] == [
        "000123", "HD-MOI", "HD-LL", None,
    ]


def test_bulk_invoice_window_uses_selected_source_rows_after_sort(
    qtbot, monkeypatch
) -> None:
    window = _window(qtbot)
    window.proxy_model.sort(
        ReviewTableModel.COLUMN_AMOUNT, Qt.SortOrder.DescendingOrder
    )
    source_row = 1
    proxy_index = window.proxy_model.mapFromSource(
        window.model.index(source_row, ReviewTableModel.COLUMN_NO)
    )
    window.table.selectRow(proxy_index.row())

    def accept_selected(dialog: BulkInvoiceDialog) -> QDialog.DialogCode:
        dialog.scope_combo.setCurrentIndex(dialog.scope_combo.findData("selected"))
        dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("overwrite"))
        dialog.new_invoice_edit.setText("HD-009")
        assert dialog.change_positions() == [source_row]
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(BulkInvoiceDialog, "exec", accept_selected)
    window.bulk_edit_invoice_numbers()

    assert [row.invoice_no for row in window.model.rows()] == [
        None, "HD-009", "HD-LL", None,
    ]
    assert window.model.dirty
    window.model.mark_clean()
    window.close()


def test_bulk_invoice_all_document_rows_include_hidden_matches(qtbot, monkeypatch) -> None:
    window = _window(qtbot)
    window.search_edit.setText("ABCD1234567")
    assert window.proxy_model.rowCount() == 1

    def accept_all(dialog: BulkInvoiceDialog) -> QDialog.DialogCode:
        dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("overwrite"))
        dialog.new_invoice_edit.setText("HD-MOI")
        assert dialog.change_positions() == [0, 1]
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(BulkInvoiceDialog, "exec", accept_all)
    window.bulk_edit_invoice_numbers()

    assert [row.invoice_no for row in window.model.rows()] == [
        "HD-MOI", "HD-MOI", "HD-LL", None,
    ]
    window.model.mark_clean()
    window.close()


def test_bulk_invoice_action_is_only_shown_for_bang_ke(qtbot) -> None:
    window = _window(qtbot)
    assert not window.bulk_invoice_button.isHidden()
    window.model.mark_clean()
    window.close()

    normal = ReviewWindow(
        {
            "metadata": {"id": 11, "status": "REVIEWING", "source_kind": "ASSISTANT"},
            "document": {"v": 4, "d": [row.to_object() for row in _rows()]},
        }
    )
    qtbot.addWidget(normal)
    assert normal.bulk_invoice_button.isHidden()
    normal.close()


def test_bulk_invoice_changes_survive_json_save(qtbot, tmp_path) -> None:
    window = _window(qtbot)
    output_path = tmp_path / "bang_ke_da_sua.json"
    codec = JsonCodec()

    def save(batch_id, document):
        assert batch_id == 10
        codec.dump_atomic(output_path, document)

    window.set_save_handler(save)
    assert window.model.bulk_update_invoice_no(
        range(window.model.rowCount()), invoice_no="0000456", mode="overwrite"
    ) == 3
    assert window.save_working()
    assert not window.model.dirty
    assert [row.invoice_no for row in codec.load(output_path).rows] == [
        "0000456", "0000456", "HD-LL", "0000456",
    ]
    window.close()
