"""Carrier choice after scanning selected BK month sheets."""

from __future__ import annotations

from copy import deepcopy
from decimal import Decimal, InvalidOperation

from PySide6.QtWidgets import (
    QAbstractItemView, QComboBox, QDialog, QDialogButtonBox, QFrame,
    QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListWidget, QMessageBox,
    QPushButton, QScrollArea, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget,
)

from app.services.excel.carrier_export import (
    CarrierExportError, CarrierExportPreview, _entered_amount,
)
from app.services.excel.payment_sync import FIELD_LABELS, NAM_FIELDS, _has_money


def _money(value: object) -> str:
    try:
        return f"{Decimal(str(value)):,.0f}".replace(",", ".") + " đ"
    except (InvalidOperation, ValueError):
        return str(value)


class CarrierRunSelectionDialog(QDialog):
    """Choose a fresh export or reopen the decisions attached to one saved run."""

    def __init__(self, runs: list[object], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Chọn lượt xuất Bên VT")
        self.resize(760, 420)
        self.resume_run_id: int | None = None
        self._runs = runs
        layout = QVBoxLayout(self)
        label = QLabel(
            "Lượt mới bắt đầu phân bổ từ đầu. Mở lại một lượt cũ sẽ giữ các lựa chọn "
            "của riêng lượt đó."
        )
        label.setWordWrap(True)
        layout.addWidget(label)
        table = QTableWidget(len(runs), 4)
        table.setHorizontalHeaderLabels(["Lượt", "Sheet BK", "Trạng thái", "Thời điểm"])
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.verticalHeader().setVisible(False)
        for index, run in enumerate(runs):
            payload = run.item_outcomes or {}
            values = (
                f"#{run.id}", ", ".join(payload.get("sheet_names", ())),
                {"SUCCEEDED": "Đã xuất", "FAILED": "Xuất lỗi",
                 "WAITING_USER": "Chờ phân bổ"}.get(run.status, run.status),
                run.completed_at or run.started_at,
            )
            for column, value in enumerate(values):
                table.setItem(index, column, QTableWidgetItem(str(value)))
        table.horizontalHeader().setStretchLastSection(True)
        table.resizeColumnsToContents()
        if runs:
            table.selectRow(0)
        self.table = table
        layout.addWidget(table)
        actions = QHBoxLayout()
        cancel = QPushButton("Hủy")
        cancel.clicked.connect(self.reject)
        fresh = QPushButton("Tạo lượt mới")
        fresh.clicked.connect(self.accept)
        reopen = QPushButton("Mở lại lượt đã chọn")
        reopen.setProperty("primary", True)
        reopen.clicked.connect(self._reopen)
        actions.addStretch(1)
        actions.addWidget(cancel)
        actions.addWidget(fresh)
        actions.addWidget(reopen)
        layout.addLayout(actions)

    def _reopen(self) -> None:
        index = self.table.currentRow()
        if 0 <= index < len(self._runs):
            self.resume_run_id = self._runs[index].id
            self.accept()


class CarrierAllocationDialog(QDialog):
    """Assign each NAM fee to the carriers named in that BK cell."""

    def __init__(self, preview: CarrierExportPreview, service: object, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Phân bổ khoản phí theo Bên VT")
        self.resize(1050, 670)
        self.preview = preview
        self.service = service
        self.resolved_preview: CarrierExportPreview | None = None
        self.rows = list(preview.allocation_rows)
        self.choices = deepcopy(dict(preview.allocation_choices))
        self._controls: dict[str, tuple[QComboBox, list[tuple[str, QLineEdit, QLineEdit | None]]]] = {}
        self._current_index = -1
        layout = QVBoxLayout(self)
        intro = QLabel(
            f"Lượt #{preview.run_id or 'mới'}: {len(self.rows)} dòng có nhiều Bên VT. "
            "Chọn bên cho từng khoản phí; nếu chia một khoản, nhập số tiền từng phần."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)
        if preview.missing_carrier_count:
            missing = QLabel(
                f"Ngoài ra có {preview.missing_carrier_count} dòng trống Bên VT; "
                "các dòng này không được xuất."
            )
            missing.setWordWrap(True)
            layout.addWidget(missing)
        middle = QHBoxLayout()
        self.row_list = QListWidget()
        self.row_list.setMinimumWidth(315)
        self.row_list.setMaximumWidth(370)
        for row in self.rows:
            self.row_list.addItem(self._row_label(row, row.allocation_complete))
        self.row_list.currentRowChanged.connect(self._change_row)
        middle.addWidget(self.row_list)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        middle.addWidget(self.scroll, 1)
        layout.addLayout(middle, 1)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Lưu phân bổ và tiếp tục")
        buttons.accepted.connect(self._accept_allocations)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.row_list.setCurrentRow(0)

    @staticmethod
    def _row_label(row: object, complete: bool) -> str:
        return (
            f"{'✓ Đã phân bổ' if complete else '○ Cần phân bổ'} · "
            f"{row.sheet_name} · QT {row.item.sqt} · {row.item.container}\n"
            + " / ".join(row.carrier_names)
        )

    def _change_row(self, index: int) -> None:
        if self._current_index >= 0:
            self._save_current()
        self._current_index = index
        if index < 0 or index >= len(self.rows):
            return
        row = self.rows[index]
        content = QWidget()
        layout = QVBoxLayout(content)
        heading = QLabel(
            f"{row.sheet_name} · QT {row.item.sqt} · {row.item.container}\n"
            f"Ô Bên VT: {row.item.carrier_value or '—'}"
        )
        heading.setWordWrap(True)
        layout.addWidget(heading)
        hint = QLabel(
            "Các tên sau chuẩn hóa: " + " · ".join(row.carrier_names)
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        quick = QHBoxLayout()
        quick_choice = QComboBox()
        for name in row.carrier_names:
            quick_choice.addItem(name, name.casefold())
        quick_button = QPushButton("Gán toàn bộ khoản phí cho bên này")
        quick_button.clicked.connect(
            lambda: self._assign_all(quick_choice.currentData())
        )
        quick.addWidget(quick_choice, 1)
        quick.addWidget(quick_button)
        layout.addLayout(quick)
        self._controls = {}
        stored = self.choices.get(row.item.item_id, {})
        for fee in NAM_FIELDS:
            amount = row.item.values.get(fee)
            if not _has_money(amount):
                continue
            card = QFrame()
            card.setFrameShape(QFrame.Shape.StyledPanel)
            group = QVBoxLayout(card)
            title = QLabel(f"{FIELD_LABELS.get(fee, fee)} · {_money(amount)}")
            group.addWidget(title)
            invoice = row.item.invoice_values.get(fee)
            if invoice:
                group.addWidget(QLabel(f"Hóa đơn nguồn: {invoice}"))
            combo = QComboBox()
            combo.addItem("Chọn bên VT nhận toàn bộ khoản này", "")
            for name in row.carrier_names:
                combo.addItem(name, name.casefold())
            combo.addItem("Chia khoản này cho nhiều bên", "__split__")
            group.addWidget(combo)
            split = QWidget()
            grid = QGridLayout(split)
            grid.addWidget(QLabel("Bên VT"), 0, 0)
            grid.addWidget(QLabel("Số tiền (đ)"), 0, 1)
            if invoice:
                grid.addWidget(QLabel("HĐ của phần này"), 0, 2)
            inputs: list[tuple[str, QLineEdit, QLineEdit | None]] = []
            for offset, name in enumerate(row.carrier_names, 1):
                amount_edit = QLineEdit()
                amount_edit.setPlaceholderText("Ví dụ 500000 hoặc 500.000")
                invoice_edit = QLineEdit() if invoice else None
                if invoice_edit:
                    invoice_edit.setPlaceholderText("Nhập hoặc xác nhận số HĐ")
                grid.addWidget(QLabel(name), offset, 0)
                grid.addWidget(amount_edit, offset, 1)
                if invoice_edit:
                    grid.addWidget(invoice_edit, offset, 2)
                inputs.append((name.casefold(), amount_edit, invoice_edit))
            group.addWidget(split)
            total_label = QLabel()
            group.addWidget(total_label)
            total_label.setVisible(False)
            def update_total(
                _text: str = "", *, fee_inputs=inputs,
                expected=amount, target=total_label,
            ) -> None:
                try:
                    total = sum(
                        (_entered_amount(edit.text()) for _, edit, _ in fee_inputs
                         if edit.text().strip()),
                        Decimal(0),
                    )
                    target.setText(
                        f"Đã phân bổ {_money(total)} / {_money(expected)}"
                        + (" · Khớp" if total == Decimal(str(expected)) else " · Chưa khớp")
                    )
                except (ValueError, ArithmeticError):
                    target.setText("Số tiền chưa hợp lệ")
            for _, amount_edit, _ in inputs:
                amount_edit.textChanged.connect(update_total)
            def toggle_split(
                _index: int, *, widget=split, total=total_label, choice=combo,
            ) -> None:
                is_split = choice.currentData() == "__split__"
                widget.setVisible(is_split)
                total.setVisible(is_split)
            combo.currentIndexChanged.connect(toggle_split)
            saved = stored.get(fee, [])
            if len(saved) > 1:
                combo.setCurrentIndex(combo.count() - 1)
                for key, amount_edit, invoice_edit in inputs:
                    match = next((entry for entry in saved if entry.get("carrier_key") == key), None)
                    if match:
                        amount_edit.setText(str(match.get("amount", "")))
                        if invoice_edit:
                            invoice_edit.setText(str(match.get("invoice", "")))
            elif len(saved) == 1:
                position = combo.findData(saved[0].get("carrier_key"))
                combo.setCurrentIndex(max(0, position))
            split.setVisible(combo.currentData() == "__split__")
            total_label.setVisible(combo.currentData() == "__split__")
            update_total()
            self._controls[fee] = combo, inputs
            layout.addWidget(card)
        layout.addStretch(1)
        self.scroll.setWidget(content)

    def _assign_all(self, carrier_key: str) -> None:
        for combo, _inputs in self._controls.values():
            index = combo.findData(carrier_key)
            if index >= 0:
                combo.setCurrentIndex(index)
        self._save_current()

    def _save_current(self) -> None:
        row = self.rows[self._current_index]
        fields: dict[str, list[dict[str, str]]] = {}
        for fee, (combo, inputs) in self._controls.items():
            choice = combo.currentData()
            if choice == "__split__":
                entries = [
                    {"carrier_key": key, "amount": amount.text().strip(),
                     "invoice": invoice.text().strip() if invoice else ""}
                    for key, amount, invoice in inputs
                    if amount.text().strip()
                ]
                if entries:
                    fields[fee] = entries
            elif choice:
                fields[fee] = [{
                    "carrier_key": choice,
                    "amount": str(row.item.values[fee]),
                    "invoice": "",
                }]
        if fields:
            self.choices[row.item.item_id] = fields
        else:
            self.choices.pop(row.item.item_id, None)
        checked = self.service.apply_allocations(
            self.preview, {row.item.item_id: fields},
            allow_partial=True, persist=False,
        )
        checked_row = next(
            item for item in checked.allocation_rows
            if item.item.item_id == row.item.item_id
        )
        self.row_list.item(self._current_index).setText(
            self._row_label(row, checked_row.allocation_complete)
        )

    def _accept_allocations(self) -> None:
        self._save_current()
        try:
            self.resolved_preview = self.service.apply_allocations(
                self.preview, self.choices
            )
        except (CarrierExportError, ValueError, ArithmeticError) as exc:
            QMessageBox.warning(self, "Phân bổ chưa hợp lệ", str(exc))
            return
        self.accept()


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
            selected = next(
                (index for index, key in enumerate(self._keys)
                 if key == preview.selected_carrier_key),
                0,
            )
            table.selectRow(selected)
        else:
            buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(False)

    def _accept_selection(self) -> None:
        if self.table.currentRow() >= 0:
            self.accept()

    @property
    def carrier_key(self) -> str | None:
        row = self.table.currentRow()
        return self._keys[row] if 0 <= row < len(self._keys) else None
