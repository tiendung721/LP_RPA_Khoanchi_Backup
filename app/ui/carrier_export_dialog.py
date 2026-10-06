"""Carrier choice after scanning selected BK month sheets."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QAbstractItemView, QDialog, QDialogButtonBox, QLabel, QTableWidget,
    QTableWidgetItem, QVBoxLayout,
)

from app.services.excel.carrier_export import CarrierExportPreview


class CarrierSelectionDialog(QDialog):
    def __init__(self, preview: CarrierExportPreview, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Chọn bên vận tải để xuất Excel")
        self.resize(720, 420)
        self._keys = [carrier.key for carrier in preview.carriers]
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            f"Đã quét {len(preview.sheet_names)} sheet BK. "
            f"Có {preview.unresolved_count} dòng chưa xác định bên VT; các dòng này không được xuất."
        ))
        table = QTableWidget(len(preview.carriers), len(preview.sheet_names) + 2)
        table.setHorizontalHeaderLabels(["Bên VT", *preview.sheet_names, "Tổng dòng"])
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.verticalHeader().setVisible(False)
        for row, carrier in enumerate(preview.carriers):
            for column, value in enumerate((carrier.name, *carrier.counts, carrier.total)):
                table.setItem(row, column, QTableWidgetItem(str(value)))
        table.horizontalHeader().setStretchLastSection(True)
        table.resizeColumnsToContents()
        layout.addWidget(table)
        self.table = table
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._accept_selection)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        if preview.carriers:
            table.selectRow(0)
        else:
            buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(False)

    def _accept_selection(self) -> None:
        if self.table.currentRow() >= 0:
            self.accept()

    @property
    def carrier_key(self) -> str | None:
        row = self.table.currentRow()
        return self._keys[row] if 0 <= row < len(self._keys) else None
