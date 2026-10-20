"""Chọn phạm vi và xem trước thao tác đổi loại cước hàng loạt."""

from __future__ import annotations

from collections import Counter
from collections.abc import Sequence

from PySide6.QtWidgets import (
    QComboBox,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QVBoxLayout,
    QWidget,
)

from .app_dialog import AppDialog
from .review_table_model import (
    FEE_CATALOG,
    ReviewRow,
    bulk_fee_change_positions,
)


class BulkFeeDialog(AppDialog):
    """Chỉ lập kế hoạch; cửa sổ review cập nhật model sau khi xác nhận."""

    def __init__(
        self,
        rows: Sequence[ReviewRow],
        *,
        selected_positions: Sequence[int] = (),
        filtered_positions: Sequence[int] = (),
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Gán/đổi loại cước hàng loạt")
        self.setMinimumWidth(480)
        self._rows = list(rows)
        self._selected_positions = set(selected_positions)
        self._filtered_positions = set(filtered_positions)

        layout = QVBoxLayout(self)
        intro = QLabel(
            "Cước biển (CB) không tham gia thao tác này. Các dòng đổi loại cước "
            "sẽ được xóa Bên vận tải; thay đổi được lưu khi bạn bấm Lưu "
            "trên màn hình kiểm tra."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        form = QFormLayout()
        self.scope_combo = QComboBox()
        self.scope_combo.setObjectName("bulkFeeScope")
        if self._selected_positions:
            self.scope_combo.addItem("Các dòng đã chọn", "selected")
        self.scope_combo.addItem("Toàn bộ dòng đang hiển thị sau lọc", "filtered")
        form.addRow("Phạm vi:", self.scope_combo)

        self.mode_combo = QComboBox()
        self.mode_combo.setObjectName("bulkFeeMode")
        self.mode_combo.addItem("Gán cho dòng chưa xác định (CXD)", "fill")
        self.mode_combo.addItem("Đổi một loại cước cũ", "replace")
        self.mode_combo.addItem("Ghi đè các dòng trong phạm vi", "overwrite")
        form.addRow("Cách cập nhật:", self.mode_combo)

        self.old_fee_combo = QComboBox()
        self.old_fee_combo.setObjectName("bulkFeeOldCode")
        form.addRow("Loại cước cũ:", self.old_fee_combo)

        self.new_fee_combo = QComboBox()
        self.new_fee_combo.setObjectName("bulkFeeNewCode")
        self.new_fee_combo.addItem("— Chọn loại cước mới —", None)
        for code, name in FEE_CATALOG.items():
            if code != "CB":
                self.new_fee_combo.addItem(f"{code} – {name}", code)
        self.new_fee_combo.setMaxVisibleItems(14)
        form.addRow("Loại cước mới:", self.new_fee_combo)
        layout.addLayout(form)
        self._form = form

        self.preview_label = QLabel()
        self.preview_label.setObjectName("bulkFeePreview")
        self.preview_label.setWordWrap(True)
        layout.addWidget(self.preview_label)
        self.existing_label = QLabel()
        self.existing_label.setObjectName("bulkFeeExistingCodes")
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

        self.scope_combo.currentIndexChanged.connect(self._refresh_old_fees)
        self.mode_combo.currentIndexChanged.connect(self._refresh_old_fees)
        self.old_fee_combo.currentIndexChanged.connect(self._update_preview)
        self.new_fee_combo.currentIndexChanged.connect(self._update_preview)
        self._refresh_old_fees()

    def scope_positions(self) -> list[int]:
        positions = (
            self._selected_positions
            if self.scope_combo.currentData() == "selected"
            else self._filtered_positions
        )
        return sorted(positions)

    @property
    def fee(self) -> str | None:
        value = self.new_fee_combo.currentData()
        return str(value) if value is not None else None

    @property
    def mode(self) -> str:
        return str(self.mode_combo.currentData() or "")

    @property
    def old_fee(self) -> str | None:
        value = self.old_fee_combo.currentData()
        return str(value) if value is not None else None

    def change_positions(self) -> list[int]:
        return bulk_fee_change_positions(
            self._rows,
            self.scope_positions(),
            fee=self.fee or "",
            mode=self.mode,
            old_fee=self.old_fee,
        )

    def _refresh_old_fees(self, *_args: object) -> None:
        previous = self.old_fee_combo.currentData()
        present = {
            fee
            for index in self.scope_positions()
            if isinstance(fee := self._rows[index].fee, str)
        }
        self.old_fee_combo.blockSignals(True)
        self.old_fee_combo.clear()
        for code, name in FEE_CATALOG.items():
            if code != "CB" and code in present:
                self.old_fee_combo.addItem(f"{code} – {name}", code)
        index = self.old_fee_combo.findData(previous)
        if index >= 0:
            self.old_fee_combo.setCurrentIndex(index)
        self.old_fee_combo.blockSignals(False)
        self._form.setRowVisible(self.old_fee_combo, self.mode == "replace")
        self._update_preview()

    def _update_preview(self, *_args: object) -> None:
        positions = self.scope_positions()
        skipped_cb = sum(self._rows[index].fee == "CB" for index in positions)
        try:
            changes = self.change_positions()
        except ValueError:
            changes = []
        cleared_carriers = sum(bool(self._rows[index].carrier) for index in changes)
        cleared_hd_rules = sum(self._rows[index].rule == "HD" for index in changes)
        preview = (
            f"Sẽ đổi {len(changes)} dòng; {skipped_cb} dòng CB bị bỏ qua; "
            f"{len(positions) - skipped_cb - len(changes)} dòng giữ nguyên. "
            f"Bên vận tải sẽ bị xóa ở {cleared_carriers} dòng."
        )
        if cleared_hd_rules:
            preview += (
                f" Quy tắc HD không hợp lệ sẽ được xóa ở {cleared_hd_rules} dòng."
            )
        self.preview_label.setText(preview)
        counts = Counter(str(self._rows[index].fee) for index in positions)
        summary = ", ".join(
            f"{code} ({count} dòng)" for code, count in counts.most_common(6)
        )
        self.existing_label.setText(
            "Loại cước hiện có: " + (summary or "không có dòng") + "."
        )
        self.apply_button.setEnabled(bool(changes))

    def accept(self) -> None:
        try:
            if self.change_positions():
                super().accept()
        except ValueError:
            return
