from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import QItemSelectionModel, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
)

from app.ui.edit_row_dialog import EditRowDialog
from app.ui.review_window import ReviewWindow, VesselVoyageNotFoundDialog
from app.ui.review_table_model import ReviewRow, ReviewTableModel
from app.ui.sea_freight_center import ReconciliationPeriodDialog
from app.database import Database
from app.repositories.batch_repository import BatchRepository
from app.sea_freight.contracts import (
    BkContainerSnapshot,
    ContainerRecord,
    VesselVoyageSuggestion,
)
from app.sea_freight.repository import SeaFreightRepository
from app.sea_freight.service import (
    SeaFreightReconciliationService,
    VesselVoyageNotFoundError,
)
from app.models import DataRow


def _review_payload() -> dict[str, Any]:
    return {
        "metadata": {
            "id": 7,
            "source_filename": "ket_qua_boc_tach.json",
            "sha256": "a" * 64,
            "status": "REVIEWING",
            "received_at": "2026-07-27T16:00:00+07:00",
        },
        "document": {
            "v": 1,
            "d": [
                [
                    "DRYU3026167",
                    None,
                    "VTN",
                    "CV",
                    "HD-130",
                    "Vận tải ABC",
                    13_554_000,
                ],
                ["VSGU2250713", "BL123456789", "CB", "HD", None, None, 27_500_000],
            ],
        },
    }


def test_cross_batch_duplicate_is_yellow_and_opens_existing_group(
    qtbot,
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "state.db")
    batches = BatchRepository(database)
    old_batch = batches.create_batch(
        source_filename="old.json",
        source_output_path=tmp_path / "old.json",
        original_archive_path=tmp_path / "old.json",
        working_path=tmp_path / "old.json",
        sha256="1" * 64,
    )
    new_batch = batches.create_batch(
        source_filename="new.json",
        source_output_path=tmp_path / "new.json",
        original_archive_path=tmp_path / "new.json",
        working_path=tmp_path / "new.json",
        sha256="2" * 64,
    )
    snapshot = BkContainerSnapshot(
        bk_path=str((tmp_path / "BK.xlsx").resolve()),
        bk_sheet="T07 26",
        vessel_voyage_raw="VIETSUN RELIANCE 2623S",
        vessel_key="VIETSUNRELIANCE",
        voyage_key="2623S",
        combined_key="VIETSUNRELIANCE2623S",
        workbook_fingerprint="bk-fp",
        snapshot_hash="snapshot",
        containers=(ContainerRecord("TSTU0000010", "T07 26", 10, 1, None),),
    )

    class Matcher:
        @staticmethod
        def snapshot(*_args: Any, **_kwargs: Any) -> BkContainerSnapshot:
            return snapshot

        @staticmethod
        def sheet_names(_path: Any) -> tuple[str, ...]:
            return ("T07 26",)

    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(repository, matcher=Matcher())
    old_row = DataRow(
        cont=None,
        bl="VSRHPG2623S06",
        fee="CB",
        rule="HD",
        amount=6_850_000,
        invoice_no="00006504",
        carrier="VIETSUN",
        vessel_voyage_raw="VIETSUN RELIANCE 2623S",
        vessel_name="VIETSUN RELIANCE",
        voyage_no="2623S",
        invoice_container_count=1,
        container_count_basis="EXPLICIT",
        invoice_date="2026-07-10",
        source_document_name="invoice.pdf",
    )
    group = service.open_or_create(
        old_row,
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=old_batch.id,
        source_item_index=0,
        source_sha256=old_batch.sha256,
    )
    repository.mark_allocated(group.id)
    repository.mark_posted(group.id, None)
    new_row = old_row.copy_with()
    new_row.source_document_name = "invoice (1).pdf"
    payload = {
        "metadata": {
            "id": new_batch.id,
            "source_filename": "new.json",
            "sha256": new_batch.sha256,
            "status": "REVIEWING",
        },
        "document": {"v": 3, "d": [new_row.to_object()]},
    }

    window = ReviewWindow(payload, sea_freight_service=service)
    qtbot.addWidget(window)
    try:
        assert window.model.validation_at(0).status.value == "warning"
        assert "hồ sơ #" in window.model.data(
            window.model.index(0, ReviewTableModel.COLUMN_MESSAGES)
        )
        assert window.model.data(
            window.model.index(0, ReviewTableModel.COLUMN_LOOKUP_ACTION)
        ) == "Xem hồ sơ"
        runtime_id = window.model.runtime_id_at(0)
        assert int(window.model.lookup_presentation(runtime_id).session_id) == group.id
        assert [item.status for item in repository.list_contribution_history(group.id)] == [
            "ACTIVE",
            "DUPLICATE",
        ]
    finally:
        window.close()
        database.close()


def test_reconciliation_period_uses_existing_bk_sheets_without_default_selection(
    qtbot,
) -> None:
    dialog = ReconciliationPeriodDialog(
        sheet_names=("Ghi chú", "T01 26", "T07 26", "T12 25"),
        month=1,
        year=2026,
    )
    qtbot.addWidget(dialog)

    assert dialog.table.rowCount() == 3
    assert [dialog.table.item(row, 0).text() for row in range(3)] == [
        "T07 26",
        "T01 26",
        "T12 25",
    ]
    assert dialog.table.currentRow() == -1
    assert not dialog.buttons.button(
        QDialogButtonBox.StandardButton.Ok
    ).isEnabled()

    dialog.table.selectRow(1)

    assert dialog.selected_sheet_name == "T01 26"
    assert dialog.period() == (1, 2026)
    assert dialog.buttons.button(QDialogButtonBox.StandardButton.Ok).isEnabled()


def test_delete_selected_row_requires_confirmation(qtbot, monkeypatch) -> None:
    window = ReviewWindow(_review_payload())
    qtbot.addWidget(window)
    window.show()
    window.table.selectRow(0)
    asked: list[str] = []

    def confirm(*args: Any, **kwargs: Any) -> QMessageBox.StandardButton:
        asked.append(str(args[2]))
        return QMessageBox.StandardButton.Yes

    monkeypatch.setattr(QMessageBox, "question", confirm)

    try:
        window.delete_selected_row()

        assert asked
        assert window.model.rowCount() == 1
        assert window.model.dirty
    finally:
        window.model.mark_clean()
        window.close()


def test_delete_multiple_selected_rows_respects_proxy_sort(qtbot, monkeypatch) -> None:
    payload = _review_payload()
    payload["document"]["d"].append(
        ["ABCD1234567", "BL-3", "VSDL", "ST", "HD-3", "Vận tải C", 850_000]
    )
    window = ReviewWindow(payload)
    qtbot.addWidget(window)
    window.show()
    window.proxy_model.sort(
        ReviewTableModel.COLUMN_AMOUNT,
        Qt.SortOrder.DescendingOrder,
    )
    expected_remaining_runtime_id = window.proxy_model.data(
        window.proxy_model.index(1, ReviewTableModel.COLUMN_NO),
        ReviewTableModel.RUNTIME_ID_ROLE,
    )
    selection = window.table.selectionModel()
    flags = (
        QItemSelectionModel.SelectionFlag.Select
        | QItemSelectionModel.SelectionFlag.Rows
    )
    selection.select(window.proxy_model.index(0, 0), flags)
    selection.select(window.proxy_model.index(2, 0), flags)
    asked: list[str] = []

    def confirm(*args: Any, **kwargs: Any) -> QMessageBox.StandardButton:
        asked.append(str(args[2]))
        return QMessageBox.StandardButton.Yes

    monkeypatch.setattr(QMessageBox, "question", confirm)

    try:
        assert window.table.selectionMode() is (
            QAbstractItemView.SelectionMode.ExtendedSelection
        )
        assert not window.edit_button.isEnabled()
        assert window.delete_button.text() == "Xóa 2 dòng"

        window.delete_selected_row()

        assert "xóa 2 dòng" in asked[0]
        assert window.model.rowCount() == 1
        assert window.model.runtime_id_at(0) == expected_remaining_runtime_id
        assert window.model.dirty
    finally:
        window.model.mark_clean()
        window.close()


def test_ctrl_s_saves_once_and_clears_dirty(qtbot, monkeypatch) -> None:
    calls: list[tuple[int, Any]] = []

    def confirm(batch_id: int, document: Any) -> None:
        calls.append((batch_id, document))

    monkeypatch.setattr(
        QMessageBox,
        "information",
        lambda *args, **kwargs: QMessageBox.StandardButton.Ok,
    )
    monkeypatch.setattr(
        QMessageBox,
        "question",
        lambda *args, **kwargs: QMessageBox.StandardButton.Yes,
    )
    window = ReviewWindow(_review_payload(), confirm_handler=confirm)
    qtbot.addWidget(window)
    window.show()
    window.model.update_row(
        0,
        ReviewRow("DRYU3026167", None, "VTN", "CV", 13_554_001),
    )
    window.activateWindow()
    window.table.setFocus()
    qtbot.wait(20)

    try:
        qtbot.keyClick(
            window.table,
            Qt.Key.Key_S,
            modifier=Qt.KeyboardModifier.ControlModifier,
        )
        qtbot.wait(50)

        assert calls
        assert calls[0][0] == 7
        assert calls[0][1].rows[0].amount == 13_554_001
        assert not window.model.dirty
    finally:
        window.model.mark_clean()
        window.close()


def test_editing_row_updates_status_and_dirty_state(qtbot) -> None:
    window = ReviewWindow(_review_payload())
    qtbot.addWidget(window)

    try:
        window.model.update_row(
            0,
            ReviewRow("DRYU3026167", None, "CB", "CV", 13_554_000),
        )

        assert window.model.dirty
        assert window.model.stats.error == 1
        assert not window.confirm_button.isEnabled()
    finally:
        window.model.mark_clean()
        window.close()


def test_review_table_uses_compact_vessel_and_reconciliation_columns(qtbot) -> None:
    window = ReviewWindow(_review_payload())
    qtbot.addWidget(window)

    try:
        hidden_columns = {
            ReviewTableModel.COLUMN_CONTAINER_COUNT_BASIS,
            ReviewTableModel.COLUMN_FEE,
            ReviewTableModel.COLUMN_RULE,
            ReviewTableModel.COLUMN_RULE_NAME,
            ReviewTableModel.COLUMN_STATUS,
        }

        assert {
            column
            for column in range(window.model.columnCount())
            if window.table.isColumnHidden(column)
        } == hidden_columns
        assert window.status_filter.isVisibleTo(window)
        assert window.status_value.text() == "Đang kiểm tra"
        assert not hasattr(window, "bk_sheet_combo")
        assert (
            window.model.headerData(
                ReviewTableModel.COLUMN_LOOKUP_ACTION,
                Qt.Orientation.Horizontal,
            )
            == ""
        )
        assert window.model.headerData(
            ReviewTableModel.COLUMN_VESSEL_VOYAGE,
            Qt.Orientation.Horizontal,
        ) == "Tàu/chuyến"
    finally:
        window.close()


def test_review_window_uses_compact_balanced_layout(qtbot) -> None:
    window = ReviewWindow(_review_payload())
    qtbot.addWidget(window)

    assert window.minimumWidth() == 920
    assert window.minimumHeight() == 600
    assert window.width() == 1120
    assert window.height() == 720

    window.show()
    qtbot.wait(20)

    try:
        assert window.add_button.geometry().top() > window.search_edit.geometry().bottom()
        assert window.table.columnWidth(ReviewTableModel.COLUMN_CONT) == 140
        assert window.table.columnWidth(ReviewTableModel.COLUMN_BL) == 130
        assert window.table.columnWidth(ReviewTableModel.COLUMN_INVOICE_NO) == 130
        assert window.table.columnWidth(ReviewTableModel.COLUMN_CARRIER) == 190
        assert window.table.columnWidth(ReviewTableModel.COLUMN_AMOUNT) == 185
        assert abs(
            window.table.columnWidth(ReviewTableModel.COLUMN_FEE_NAME)
            - window.table.columnWidth(ReviewTableModel.COLUMN_MESSAGES)
        ) <= 1
    finally:
        window.close()


def test_edit_dialog_hides_rule_but_preserves_its_value(qtbot) -> None:
    dialog = EditRowDialog(
        ReviewRow(
            "DRYU3026167",
            None,
            "VTN",
            "CV",
            13_554_000,
            invoice_no="HD-130",
            carrier="Vận tải ABC",
        )
    )
    qtbot.addWidget(dialog)
    dialog.show()

    try:
        assert not dialog.rule_combo.isVisible()
        assert not hasattr(dialog, "amount_unknown")

        dialog.amount_edit.setText("13.555.000")
        edited = dialog.row_data()

        assert edited.rule == "CV"
        assert edited.amount == 13_555_000
        assert edited.invoice_no == "HD-130"
        assert edited.carrier == "Vận tải ABC"

        dialog.amount_edit.clear()
        assert dialog.row_data().amount is None
        assert any(
            "Số tiền chưa xác định" in message
            for message in dialog._last_validation.warnings
        )
    finally:
        dialog.close()


def test_daily_sync_fee_keeps_carrier_editable_and_clears_stale_fee_carrier(
    qtbot,
) -> None:
    dialog = EditRowDialog(
        ReviewRow(
            "DRYU3026167",
            None,
            "NH",
            "CV",
            13_554_000,
            carrier="NHÀ CUNG CẤP HĐ PHÍ",
        )
    )
    qtbot.addWidget(dialog)

    dialog.fee_combo.setCurrentIndex(dialog.fee_combo.findData("VTN"))

    assert dialog.carrier_edit.isEnabled()
    assert dialog.carrier_edit.text() == ""
    dialog.carrier_edit.setText("USER CHỌN")
    assert dialog.row_data().carrier == "USER CHỌN"


def test_edit_dialog_preserves_ai_raw_and_previews_effective_vessel_voyage(
    qtbot,
) -> None:
    dialog = EditRowDialog(
        ReviewRow(
            cont=None,
            bl="BL-1",
            fee="CB",
            rule="HD",
            amount=100,
            vessel_voyage_raw="NEW VISON 2610S",
            vessel_name="NEW VISON",
            voyage_no="2610S",
            invoice_container_count=1,
            container_count_basis="EXPLICIT",
        )
    )
    qtbot.addWidget(dialog)

    dialog.vessel_name_edit.setText("NEW VISION")

    assert dialog.vessel_voyage_preview.text() == "NEW VISION 2610S"
    assert dialog.row_data().vessel_voyage_raw == "NEW VISON 2610S"
    assert dialog.row_data().vessel_name == "NEW VISION"
    assert not hasattr(dialog, "vessel_voyage_raw_edit")


def test_vessel_not_found_dialog_lists_read_only_suggestions(qtbot) -> None:
    dialog = VesselVoyageNotFoundDialog(
        VesselVoyageNotFoundError(
            vessel_voyage="NEW VISON 2610S",
            bk_sheet="T07 26",
            suggestions=(
                VesselVoyageSuggestion("NEW VISION 2610S", 2, 0.95),
                VesselVoyageSuggestion("NEW VISION 2611N", 1, 0.80),
            ),
        )
    )
    qtbot.addWidget(dialog)

    suggestion_text = dialog.findChild(QLabel, "vesselVoyageSuggestions")
    assert suggestion_text is not None
    assert "NEW VISION 2610S — 2 container" in suggestion_text.text()
    assert dialog.edit_button.text() == "Sửa dòng"


def test_add_dialog_keeps_rule_selector(qtbot) -> None:
    dialog = EditRowDialog()
    qtbot.addWidget(dialog)
    dialog.show()

    try:
        assert dialog.rule_combo.isVisible()
        assert not hasattr(dialog, "amount_unknown")
    finally:
        dialog.close()


def test_save_button_writes_once_and_shows_simple_success_message(
    qtbot, monkeypatch
) -> None:
    confirmed: list[tuple[int, Any]] = []
    messages: list[tuple[str, str]] = []

    def confirm(batch_id: int, document: Any) -> None:
        confirmed.append((batch_id, document))

    def information(*args: Any, **kwargs: Any) -> QMessageBox.StandardButton:
        messages.append((str(args[1]), str(args[2])))
        return QMessageBox.StandardButton.Ok

    monkeypatch.setattr(
        QMessageBox,
        "information",
        information,
    )
    monkeypatch.setattr(
        QMessageBox,
        "question",
        lambda *args, **kwargs: QMessageBox.StandardButton.Yes,
    )
    window = ReviewWindow(
        _review_payload(),
        confirm_handler=confirm,
    )
    qtbot.addWidget(window)
    window.model.update_row(
        0,
        ReviewRow("DRYU3026167", None, "VTN", "CV", 13_554_001),
    )
    window.show()

    try:
        assert window.confirm_button.text() == "Lưu"
        assert not hasattr(window, "save_button")

        qtbot.mouseClick(window.confirm_button, Qt.MouseButton.LeftButton)

        assert len(confirmed) == 1
        assert confirmed[0][0] == 7
        assert confirmed[0][1].rows[0].amount == 13_554_001
        assert messages == [("Lưu thành công", "Đã lưu thành công.")]
        assert window.status_value.text() == "Đã xác nhận"
        assert not window.model.dirty
    finally:
        window.model.mark_clean()
        window.close()
