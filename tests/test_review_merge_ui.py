from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from PySide6.QtWidgets import QDialog, QDialogButtonBox

from app.config import AppSettings
from app.services.batch_service import BatchService
from app.ui.edit_row_dialog import EditRowDialog
from app.ui.review_merge_dialog import ReviewMergeDialog
from app.ui.review_table_model import ReviewRow
from app.ui.review_window import ReviewWindow


def _row(amount: int, *, fee: str = "VSDL", invoice_no: str = "26955") -> ReviewRow:
    return ReviewRow(
        cont="DRYU3026167",
        fee=fee,
        rule="ST",
        amount=amount,
        invoice_no=invoice_no,
        invoice_date="2026-08-29",
        source_document_id="DOC_001",
        source_document_name="Hoa_don.pdf",
    )


def _payload(rows: list[ReviewRow], *, source_kind: str = "ASSISTANT", saved: bool = False) -> dict[str, Any]:
    return {
        "metadata": {
            "id": 15,
            "status": "REVIEWING",
            "source_kind": source_kind,
            "last_saved_at": "2026-08-29T12:00:00+07:00" if saved else None,
        },
        "document": {"v": 4, "d": [row.to_object() for row in rows]},
    }


def test_open_gpt_review_merges_then_save_serializes_one_total(qtbot) -> None:
    originals = [_row(291_600), _row(70_200), _row(75_600)]
    saved_documents: list[Any] = []
    window = ReviewWindow(
        _payload(originals),
        save_handler=lambda _batch_id, document: saved_documents.append(document),
    )
    qtbot.addWidget(window)
    try:
        assert window.model.rowCount() == 1
        assert window.model.row_at(0).amount == 437_400
        assert window.model.row_at(0).fee == "VSDL"
        assert window.model.row_at(0).rule == "ST"
        assert window.model.dirty
        assert window._pending_deleted_source_indices == {1, 2}
        assert window.save_working()
        assert len(saved_documents[0].rows) == 1
        assert saved_documents[0].rows[0].amount == 437_400
        assert not window.model.dirty

        reopened = ReviewWindow(_payload([window.model.row_at(0)], saved=True))
        qtbot.addWidget(reopened)
        assert reopened.model.rowCount() == 1
        assert reopened.model.stats.total_amount == 437_400
        assert not reopened.model.dirty
        reopened.close()
    finally:
        window.model.mark_clean()
        window.close()


def test_service_saves_and_reopens_merged_json_once(qtbot, tmp_path: Path) -> None:
    settings = AppSettings(data_root=tmp_path / "runtime", output_dir=tmp_path / "Output")
    source = settings.output_dir / "ket_qua_boc_tach.json"
    source.parent.mkdir(parents=True)
    source.write_text(
        json.dumps(
            {"v": 4, "d": [
                _row(291_600).to_object(),
                _row(70_200).to_object(),
                _row(75_600).to_object(),
            ]},
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    service = BatchService(settings)
    window = None
    reopened = None
    try:
        received = service.receive_file(source)
        assert received.review is not None
        window = ReviewWindow(received.review, batch_service=service)
        qtbot.addWidget(window)
        assert window.model.rowCount() == 1
        assert window.save_working()

        persisted = json.loads(received.batch.working_path.read_text(encoding="utf-8"))
        assert persisted["v"] == 4
        assert len(persisted["d"]) == 1
        assert persisted["d"][0]["amount"] == 437_400

        reopened = ReviewWindow(service.load_batch(received.batch.id), batch_service=service)
        qtbot.addWidget(reopened)
        assert reopened.model.rowCount() == 1
        assert reopened.model.row_at(0).amount == 437_400
        assert not reopened.model.dirty
    finally:
        if reopened is not None:
            reopened.close()
        if window is not None:
            window.model.mark_clean()
            window.close()
        service.close()


def test_editing_fee_merges_with_existing_rows(qtbot, monkeypatch) -> None:
    window = ReviewWindow(_payload([_row(291_600, fee="HV"), _row(70_200), _row(75_600, fee="CXD")]))
    qtbot.addWidget(window)
    original_runtime = window.model.runtime_id_at(0)
    monkeypatch.setattr(EditRowDialog, "exec", lambda _dialog: QDialog.DialogCode.Accepted)
    monkeypatch.setattr(
        EditRowDialog,
        "row_data",
        lambda dialog: ReviewRow.from_mapping(
            {**dialog._collect_row()[0].to_object(), "fee": "VSDL"}
        ),
    )
    try:
        assert window.model.rowCount() == 3
        assert window._edit_source_row(0)
        assert window.model.rowCount() == 2
        assert window.model.runtime_id_at(0) == original_runtime
        assert window._edit_source_row(1)
        assert window.model.rowCount() == 1
        assert window.model.row_at(0).amount == 437_400
        assert window.model.row_at(0).fee == "VSDL"
        assert window._pending_deleted_source_indices == {1, 2}
    finally:
        window.model.mark_clean()
        window.close()


def test_conflicting_invoice_requires_selection_and_cancel_keeps_rows(qtbot, monkeypatch) -> None:
    window = ReviewWindow(_payload([_row(100, invoice_no="INV-A"), _row(200, invoice_no="INV-B")]))
    qtbot.addWidget(window)
    try:
        assert window.model.rowCount() == 2
        monkeypatch.setattr(ReviewMergeDialog, "exec", lambda _dialog: QDialog.DialogCode.Rejected)
        window._run_merge_review(show_dialogs=True)
        assert window.model.rowCount() == 2
        assert not window.model.dirty
    finally:
        window.close()

    window = ReviewWindow(_payload([_row(100, invoice_no="INV-A"), _row(200, invoice_no="INV-B")]))
    qtbot.addWidget(window)
    monkeypatch.setattr(ReviewMergeDialog, "exec", lambda _dialog: QDialog.DialogCode.Accepted)
    monkeypatch.setattr(ReviewMergeDialog, "selected_values", lambda _dialog: {"invoice_no": "INV-B"})
    try:
        window._run_merge_review(show_dialogs=True)
        assert window.model.rowCount() == 1
        assert window.model.row_at(0).invoice_no == "INV-B"
        assert window.model.row_at(0).amount == 300
    finally:
        window.model.mark_clean()
        window.close()


def test_accepting_conflict_does_not_merge_an_unmatched_third_line(qtbot, monkeypatch) -> None:
    first = _row(100, invoice_no="A")
    first.invoice_date = "2026-08-28"
    second = _row(200, invoice_no="A")
    third = _row(300, invoice_no="B")
    window = ReviewWindow(_payload([first, second, third]))
    qtbot.addWidget(window)
    monkeypatch.setattr(ReviewMergeDialog, "exec", lambda _dialog: QDialog.DialogCode.Accepted)
    monkeypatch.setattr(
        ReviewMergeDialog,
        "selected_values",
        lambda _dialog: {"invoice_date": "2026-08-29"},
    )
    try:
        window._run_merge_review(show_dialogs=True)
        assert window.model.rowCount() == 2
        assert [row.amount for row in window.model.rows()] == [300, 300]
        window._run_merge_review(show_dialogs=True)
        assert window.model.rowCount() == 2
    finally:
        window.model.mark_clean()
        window.close()


def test_review_dialog_requires_explicit_conflict_choice(qtbot) -> None:
    dialog = ReviewMergeDialog(
        [_row(100, invoice_no="INV-A"), _row(200, invoice_no="INV-B")],
        [0, 1],
        ("invoice_no",),
    )
    qtbot.addWidget(dialog)
    button = dialog.buttons.button(QDialogButtonBox.StandardButton.Ok)
    assert not button.isEnabled()
    dialog._selectors["invoice_no"].setCurrentIndex(2)
    assert button.isEnabled()
    assert dialog.selected_values() == {"invoice_no": "INV-B"}
    dialog.close()


def test_bang_ke_source_is_not_merged(qtbot) -> None:
    window = ReviewWindow(_payload([_row(100), _row(200)], source_kind="BANG_KE"))
    qtbot.addWidget(window)
    assert window.model.rowCount() == 2
    window.close()


def test_cb_cxd_exact_duplicate_and_missing_amount_are_not_auto_merged(qtbot) -> None:
    cb_rows = [_row(100, fee="CB"), _row(200, fee="CB")]
    for row in cb_rows:
        row.rule = "HD"
    cases = [
        cb_rows,
        [_row(100, fee="CXD"), _row(200, fee="CXD")],
        [_row(100), _row(100)],
        [_row(100), _row(200)],
        [_row(100), _row(200)],
    ]
    cases[3][1].amount = None
    cases[4][1].carrier = ["không hợp lệ"]
    for rows in cases:
        window = ReviewWindow(_payload(rows))
        qtbot.addWidget(window)
        assert window.model.rowCount() == 2
        assert not window.model.dirty
        window.close()


def test_missing_amount_does_not_block_other_valid_lines(qtbot) -> None:
    missing = _row(100)
    missing.amount = None
    window = ReviewWindow(_payload([missing, _row(200), _row(300)]))
    qtbot.addWidget(window)
    try:
        assert window.model.rowCount() == 2
        assert [row.amount for row in window.model.rows()] == [None, 500]
    finally:
        window.model.mark_clean()
        window.close()


def test_fee_change_exposes_rule_only_when_needed(qtbot) -> None:
    dialog = EditRowDialog(_row(100))
    qtbot.addWidget(dialog)
    dialog.show()
    try:
        assert not dialog.rule_combo.isVisible()
        dialog.fee_combo.setCurrentIndex(dialog.fee_combo.findData("CB"))
        assert dialog.rule_combo.isVisible()
        assert dialog._collect_row()[0].rule == "HD"
        dialog.fee_combo.setCurrentIndex(dialog.fee_combo.findData("VSDL"))
        assert not dialog.rule_combo.isVisible()
        assert dialog._collect_row()[0].rule == "ST"
    finally:
        dialog.close()


def test_merge_reindexes_source_links_only_after_save(qtbot) -> None:
    sync_calls: list[dict[str, Any]] = []
    sea_service = SimpleNamespace(
        apply_batch_row_deletions=lambda **kwargs: sync_calls.append(kwargs),
        sync_batch_history=lambda *_args, **_kwargs: {},
        repository=SimpleNamespace(groups_for_source_batch=lambda _batch_id: {}),
    )
    sea_row = _row(300, fee="CB")
    sea_row.cont = None
    sea_row.rule = "HD"
    window = ReviewWindow(
        _payload([_row(100), _row(200), sea_row]),
        sea_freight_service=sea_service,
        save_handler=lambda *_args: None,
    )
    qtbot.addWidget(window)
    try:
        assert window.model.rowCount() == 2
        assert sync_calls == []
        assert window.save_working()
        assert sync_calls == [{
            "source_batch_id": 15,
            "deleted_source_indices": {1},
            "remaining_source_indices": {0: 0, 2: 1},
        }]
    finally:
        window.model.mark_clean()
        window.close()
