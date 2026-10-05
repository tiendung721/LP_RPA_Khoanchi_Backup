"""Chọn phạm vi và xem trước thao tác gán số HĐ cho bảng kê nguồn."""

from __future__ import annotations

from collections import Counter
from collections.abc import Sequence

from PySide6.QtWidgets import (
    QComboBox,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QVBoxLayout,
    QWidget,
)

from .app_dialog import AppDialog
from .review_table_model import ReviewRow, bulk_invoice_change_positions


class BulkInvoiceDialog(AppDialog):
    """Chỉ lập kế hoạch; dữ liệu được cập nhật sau khi hộp thoại đóng thành công."""

    def __init__(
        self,
        rows: Sequence[ReviewRow],
        *,
        selected_positions: Sequence[int] = (),
        document_id: str | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Gán/đổi số HĐ hàng loạt")
        self.setMinimumWidth(440)
        self._rows = list(rows)
        self._selected_positions = set(selected_positions)

        layout = QVBoxLayout(self)
        intro = QLabel(
            "Chọn các dòng cần cập nhật. Thay đổi sẽ được lưu khi bạn bấm Lưu "
            "trên màn hình kiểm tra."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        form = QFormLayout()
        self.document_combo = QComboBox()
        self.document_combo.setObjectName("bulkInvoiceDocument")
        documents: dict[str, tuple[str, int]] = {}
        for row in self._rows:
            name, count = documents.get(row.source_document_id, (row.source_document_name, 0))
            documents[row.source_document_id] = (name, count + 1)
        for source_id, (name, count) in documents.items():
            self.document_combo.addItem(f"{name} – {count} dòng", source_id)
        preferred_index = self.document_combo.findData(document_id)
        if preferred_index >= 0:
            self.document_combo.setCurrentIndex(preferred_index)
        form.addRow("Bảng kê nguồn:", self.document_combo)

        self.scope_combo = QComboBox()
        self.scope_combo.setObjectName("bulkInvoiceScope")
        self.scope_combo.addItem("Toàn bộ bảng kê", "all")
        if self._selected_positions:
            self.scope_combo.addItem("Các dòng đã chọn", "selected")
            self.scope_combo.setCurrentIndex(self.scope_combo.findData("selected"))
        form.addRow("Phạm vi:", self.scope_combo)

        self.mode_combo = QComboBox()
        self.mode_combo.setObjectName("bulkInvoiceMode")
        self.mode_combo.addItem("Điền dòng chưa có số HĐ", "fill")
        self.mode_combo.addItem("Đổi một số HĐ cũ", "replace")
        self.mode_combo.addItem("Ghi đè tất cả", "overwrite")
        form.addRow("Cách cập nhật:", self.mode_combo)

        self.old_invoice_combo = QComboBox()
        self.old_invoice_combo.setObjectName("bulkInvoiceOldNumber")
        form.addRow("Số HĐ cũ:", self.old_invoice_combo)

        self.new_invoice_edit = QLineEdit()
        self.new_invoice_edit.setObjectName("bulkInvoiceNewNumber")
        self.new_invoice_edit.setMaxLength(200)
        self.new_invoice_edit.setPlaceholderText("Nhập số HĐ mới")
        form.addRow("Số HĐ mới:", self.new_invoice_edit)
        layout.addLayout(form)
        self._form = form

        self.preview_label = QLabel()
        self.preview_label.setObjectName("bulkInvoicePreview")
        self.preview_label.setWordWrap(True)
        layout.addWidget(self.preview_label)
        self.existing_label = QLabel()
        self.existing_label.setObjectName("bulkInvoiceExistingNumbers")
        self.existing_label.setWordWrap(True)
        layout.addWidget(self.existing_label)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.apply_button = buttons.button(QDialogButtonBox.StandardButton.Ok)
        self.apply_button.setText("Áp dụng")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Hủy")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.document_combo.currentIndexChanged.connect(self._refresh_old_invoices)
        self.scope_combo.currentIndexChanged.connect(self._refresh_old_invoices)
        self.mode_combo.currentIndexChanged.connect(self._refresh_old_invoices)
        self.old_invoice_combo.currentIndexChanged.connect(self._update_preview)
        self.new_invoice_edit.textChanged.connect(self._update_preview)
        self._refresh_old_invoices()

    def scope_positions(self) -> list[int]:
        source_id = self.document_combo.currentData()
        selected_only = self.scope_combo.currentData() == "selected"
        return [
            index
            for index, row in enumerate(self._rows)
            if row.source_document_id == source_id
            and (not selected_only or index in self._selected_positions)
        ]

    @property
    def invoice_no(self) -> str:
        return " ".join(self.new_invoice_edit.text().split())

    @property
    def mode(self) -> str:
        return str(self.mode_combo.currentData() or "")

    @property
    def old_invoice_no(self) -> str | None:
        value = self.old_invoice_combo.currentData()
        return str(value) if value is not None else None

    def change_positions(self) -> list[int]:
        return bulk_invoice_change_positions(
            self._rows,
            self.scope_positions(),
            invoice_no=self.invoice_no,
            mode=self.mode,
            old_invoice_no=self.old_invoice_no,
        )

    def _refresh_old_invoices(self, *_args: object) -> None:
        previous = self.old_invoice_combo.currentData()
        self.old_invoice_combo.blockSignals(True)
        self.old_invoice_combo.clear()
        seen: set[str] = set()
        for position in self.scope_positions():
            row = self._rows[position]
            if row.fee == "LL" or not isinstance(row.invoice_no, str):
                continue
            value = " ".join(row.invoice_no.split())
            key = value.casefold()
            if value and key not in seen:
                seen.add(key)
                self.old_invoice_combo.addItem(value, value)
        previous_index = self.old_invoice_combo.findData(previous)
        if previous_index >= 0:
            self.old_invoice_combo.setCurrentIndex(previous_index)
        self.old_invoice_combo.blockSignals(False)
        self._form.setRowVisible(self.old_invoice_combo, self.mode == "replace")
        self._update_preview()

    def _update_preview(self, *_args: object) -> None:
        eligible_rows = [
            self._rows[position]
            for position in self.scope_positions()
            if self._rows[position].fee != "LL"
        ]
        eligible_count = len(eligible_rows)
        try:
            change_count = len(self.change_positions())
        except ValueError:
            change_count = 0
        self.preview_label.setText(
            f"Sẽ cập nhật {change_count} dòng; "
            f"{eligible_count - change_count} dòng giữ nguyên."
        )
        existing = Counter(
            " ".join(row.invoice_no.split())
            for row in eligible_rows
            if isinstance(row.invoice_no, str) and row.invoice_no.strip()
        )
        if existing:
            entries = [
                f"{value} ({count} dòng)"
                for value, count in existing.most_common(5)
            ]
            if len(existing) > 5:
                entries.append(f"và {len(existing) - 5} số khác")
            self.existing_label.setText("Số HĐ hiện có: " + ", ".join(entries) + ".")
        else:
            self.existing_label.setText("Số HĐ hiện có: chưa có.")
        self.apply_button.setEnabled(change_count > 0)

    def accept(self) -> None:
        try:
            changes = self.change_positions()
        except ValueError:
            return
        if changes:
            super().accept()
