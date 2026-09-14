"""Hộp thoại xem và chọn SQT trước khi gọi PAD."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any, Callable, Iterable

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QComboBox,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.rpa_expense.contracts import (
    RPA_STATUS_IMPORTED,
    RPA_STATUS_NOT_IMPORTED,
)

from .app_dialog import AppDialog


STATUS_ROLE = Qt.ItemDataRole.UserRole + 1


class _SortableTableWidgetItem(QTableWidgetItem):
    """QTableWidgetItem hiển thị text nhưng sắp theo giá trị nghiệp vụ."""

    def __init__(self, text: str, sort_value: Any) -> None:
        super().__init__(text)
        self.sort_value = sort_value

    def __lt__(self, other: QTableWidgetItem) -> bool:
        if isinstance(other, _SortableTableWidgetItem):
            return self.sort_value < other.sort_value
        return super().__lt__(other)


def _money(value: Any) -> str:
    try:
        return f"{int(value or 0):,}".replace(",", ".")
    except (TypeError, ValueError):
        return str(value or 0)


class RpaSqtSelectionDialog(AppDialog):
    COLUMNS = (
        "Chọn",
        "SQT",
        "Trạng thái RPA",
        "Dòng BK",
        "Cước MB",
        "N.hạ MB",
        "Cước biển",
        "N.hạ/VS/D/O/Lệnh",
        "Cước MN",
        "Lưu cont/Quá tải",
        "Sửa chữa",
        "Tổng",
        "Kiểm tra",
    )

    def __init__(
        self,
        plan: Any,
        parent: QWidget | None = None,
        *,
        initial_selected_sqt: Iterable[str] = (),
        restore_info: Mapping[str, Any] | None = None,
        clear_saved_callback: Callable[[], bool] | None = None,
    ) -> None:
        super().__init__(parent)
        self.plan = plan
        self._initial_selected_sqt = {
            str(value) for value in initial_selected_sqt
        }
        self.restore_info = dict(restore_info or {})
        self._skipped_sqt = {
            str(value) for value in self.restore_info.get("skipped_sqt", ())
        }
        self._clear_saved_callback = clear_saved_callback
        self._checks: list[QTableWidgetItem] = []
        self._active_filter: str | None = None
        self._sort_options = {
            "status_asc": (2, Qt.SortOrder.AscendingOrder),
            "status_desc": (2, Qt.SortOrder.DescendingOrder),
            "sqt_asc": (1, Qt.SortOrder.AscendingOrder),
            "sqt_desc": (1, Qt.SortOrder.DescendingOrder),
            "total_desc": (11, Qt.SortOrder.DescendingOrder),
            "total_asc": (11, Qt.SortOrder.AscendingOrder),
            "rows_asc": (3, Qt.SortOrder.AscendingOrder),
            "rows_desc": (3, Qt.SortOrder.DescendingOrder),
        }
        self.setObjectName("rpaSqtSelectionDialog")
        self.setWindowTitle("Chọn số quyết toán chạy RPA")
        self.resize(1380, 650)
        self.setStyleSheet(
            """
            QPushButton[filterChip="true"] {
                padding: 5px 12px;
                border-radius: 13px;
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                color: #475569;
                font-weight: 600;
            }
            QPushButton[filterChip="true"]:checked {
                background: #E8F1FF;
                border-color: #2563EB;
                color: #1D4ED8;
            }
            QLabel#rpaSelectionSummaryLabel {
                color: #334155;
                font-weight: 700;
            }
            """
        )
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        title = QLabel(
            f"Sheet {self.plan.sheet_name}: chọn các SQT cần nhập lên phần mềm quyết toán."
        )
        title.setStyleSheet("font-size: 12pt; font-weight: 700;")
        title.setWordWrap(True)
        layout.addWidget(title)
        note = QLabel(
            "SQT “Đã nhập” vẫn được phép chọn và chạy lại. "
            "Trạng thái chỉ được đổi sang “Đã nhập” sau khi PAD xác nhận đã lưu thành công trên web."
        )
        note.setWordWrap(True)
        note.setProperty("muted", True)
        layout.addWidget(note)

        if self.restore_info.get("found"):
            restored = int(self.restore_info.get("restored_count", 0) or 0)
            saved = int(self.restore_info.get("saved_count", 0) or 0)
            skipped = max(0, saved - restored)
            restore_row = QHBoxLayout()
            message = f"Đã khôi phục {restored}/{saved} SQT từ lần chạy gần nhất."
            if skipped:
                message += f" Có {skipped} SQT đã thay đổi hoặc không còn hợp lệ."
            self.restore_label = QLabel(message)
            self.restore_label.setWordWrap(True)
            self.restore_label.setProperty("status", "info")
            restore_row.addWidget(self.restore_label, 1)
            if self._clear_saved_callback is not None:
                clear_saved = QPushButton("Xóa lựa chọn đã nhớ cho file này")
                clear_saved.setObjectName("clearRememberedRpaSelectionButton")
                clear_saved.clicked.connect(self._clear_remembered)
                restore_row.addWidget(clear_saved)
            layout.addLayout(restore_row)

        filters = QHBoxLayout()
        filters.setSpacing(7)
        filters.addWidget(QLabel("Hiển thị:"))
        self.filter_group = QButtonGroup(self)
        self.filter_group.setExclusive(True)
        imported_count = sum(
            item.status == RPA_STATUS_IMPORTED for item in self.plan.items
        )
        filter_specs = (
            ("Tất cả", None, len(self.plan.items), "allRpaFilterButton"),
            (
                "Chưa nhập",
                RPA_STATUS_NOT_IMPORTED,
                len(self.plan.items) - imported_count,
                "notImportedRpaFilterButton",
            ),
            (
                "Đã nhập",
                RPA_STATUS_IMPORTED,
                imported_count,
                "importedRpaFilterButton",
            ),
        )
        self.filter_buttons: dict[str | None, QPushButton] = {}
        for label, status, count, object_name in filter_specs:
            button = QPushButton(f"{label} ({count})")
            button.setObjectName(object_name)
            button.setProperty("filterChip", True)
            button.setCheckable(True)
            button.clicked.connect(
                lambda checked, value=status: checked and self._set_filter(value)
            )
            self.filter_group.addButton(button)
            self.filter_buttons[status] = button
            filters.addWidget(button)
        self.filter_buttons[None].setChecked(True)
        filters.addStretch()
        filters.addWidget(QLabel("Sắp xếp:"))
        self.sort_combo = QComboBox()
        self.sort_combo.setObjectName("rpaSortCombo")
        self.sort_combo.addItem("Trạng thái: Chưa nhập trước", "status_asc")
        self.sort_combo.addItem("Trạng thái: Đã nhập trước", "status_desc")
        self.sort_combo.addItem("SQT tăng dần", "sqt_asc")
        self.sort_combo.addItem("SQT giảm dần", "sqt_desc")
        self.sort_combo.addItem("Tổng tiền giảm dần", "total_desc")
        self.sort_combo.addItem("Tổng tiền tăng dần", "total_asc")
        self.sort_combo.addItem("Dòng BK tăng dần", "rows_asc")
        self.sort_combo.addItem("Dòng BK giảm dần", "rows_desc")
        self.sort_combo.currentIndexChanged.connect(self._apply_sort)
        filters.addWidget(self.sort_combo)
        layout.addLayout(filters)

        selection = QHBoxLayout()
        select_visible = QPushButton("Chọn các SQT đang hiển thị")
        clear_visible = QPushButton("Bỏ chọn đang hiển thị")
        select_visible.setObjectName("selectVisibleRpaSqtButton")
        clear_visible.setObjectName("clearVisibleRpaSqtButton")
        select_visible.clicked.connect(
            lambda: self._set_visible(Qt.CheckState.Checked)
        )
        clear_visible.clicked.connect(
            lambda: self._set_visible(Qt.CheckState.Unchecked)
        )
        select_all = QPushButton("Chọn tất cả có thể chạy")
        clear_all = QPushButton("Bỏ chọn tất cả")
        select_all.setObjectName("selectAllRpaSqtButton")
        clear_all.setObjectName("clearAllRpaSqtButton")
        select_all.clicked.connect(
            lambda: self._set_all(Qt.CheckState.Checked)
        )
        clear_all.clicked.connect(
            lambda: self._set_all(Qt.CheckState.Unchecked)
        )
        selection.addWidget(select_visible)
        selection.addWidget(clear_visible)
        selection.addWidget(select_all)
        selection.addWidget(clear_all)
        selection.addStretch()
        self.selection_summary = QLabel()
        self.selection_summary.setObjectName("rpaSelectionSummaryLabel")
        selection.addWidget(self.selection_summary)
        layout.addLayout(selection)

        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setObjectName("rpaSqtTable")
        self.table.setHorizontalHeaderLabels(list(self.COLUMNS))
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.table.horizontalHeader().setSectionResizeMode(
            12, QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self.table, 1)

        for item in self.plan.items:
            self._add_item(item)

        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().sortIndicatorChanged.connect(
            lambda *_: QTimer.singleShot(0, self._apply_filter)
        )
        self.table.itemChanged.connect(self._on_item_changed)
        self._apply_sort()
        self._update_selection_summary()

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.setObjectName("rpaSqtDialogButtons")
        self.run_button = buttons.button(QDialogButtonBox.StandardButton.Ok)
        self.run_button.setText("Chạy RPA các SQT đã chọn")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Hủy")
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self._update_selection_summary()

    def _add_item(self, source: Any) -> None:
        row = self.table.rowCount()
        self.table.insertRow(row)
        check = QTableWidgetItem()
        flags = Qt.ItemFlag.ItemIsSelectable
        if source.can_run:
            flags |= Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable
            check.setCheckState(
                Qt.CheckState.Checked
                if str(source.sqt) in self._initial_selected_sqt
                else Qt.CheckState.Unchecked
            )
        check.setFlags(flags)
        check.setData(Qt.ItemDataRole.UserRole, source.sqt)
        check.setData(STATUS_ROLE, source.status)
        check.setToolTip(
            "Đã được PAD xác nhận lưu thành công; vẫn có thể chọn để nhập lại."
            if source.status == RPA_STATUS_IMPORTED
            else "Chưa có xác nhận lưu thành công từ PAD."
        )
        if str(source.sqt) in self._skipped_sqt:
            check.setBackground(QColor("#FFF4CC"))
        self._checks.append(check)
        self.table.setItem(row, 0, check)

        amounts = source.amounts
        status_display = (
            "✓ Đã nhập"
            if source.status == RPA_STATUS_IMPORTED
            else "● Chưa nhập"
        )
        sqt_sort = self._sqt_sort_key(source.sqt)
        source_rows = tuple(int(value) for value in source.source_rows)
        values = (
            source.sqt,
            status_display,
            ", ".join(str(value) for value in source.source_rows),
            _money(amounts.cuoc_bo_dong_hang),
            _money(amounts.nang_ha_dong_hang),
            _money(amounts.cuoc_bien),
            _money(amounts.nang_do_vs_lam_lenh),
            _money(amounts.cuoc_bo_tra_hang),
            _money(amounts.luu_cont_qua_tai),
            _money(amounts.sua_chua_cont),
            _money(amounts.total),
            source.validation_message or "Hợp lệ",
        )
        sort_values = (
            sqt_sort,
            (
                1 if source.status == RPA_STATUS_IMPORTED else 0,
                sqt_sort,
            ),
            source_rows,
            amounts.cuoc_bo_dong_hang,
            amounts.nang_ha_dong_hang,
            amounts.cuoc_bien,
            amounts.nang_do_vs_lam_lenh,
            amounts.cuoc_bo_tra_hang,
            amounts.luu_cont_qua_tai,
            amounts.sua_chua_cont,
            amounts.total,
            source.validation_message or "Hợp lệ",
        )
        for column, (value, sort_value) in enumerate(
            zip(values, sort_values), 1
        ):
            cell = _SortableTableWidgetItem(str(value), sort_value)
            cell.setToolTip(str(value))
            if column == 2:
                cell.setData(STATUS_ROLE, source.status)
                font = cell.font()
                font.setBold(True)
                cell.setFont(font)
                cell.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if source.status == RPA_STATUS_IMPORTED:
                    cell.setForeground(QColor("#15803D"))
                    cell.setBackground(QColor("#ECFDF3"))
                    cell.setToolTip(
                        "Đã được PAD xác nhận lưu thành công; "
                        "vẫn có thể chọn để nhập lại."
                    )
                else:
                    cell.setForeground(QColor("#1D4ED8"))
                    cell.setBackground(QColor("#EFF6FF"))
                    cell.setToolTip("Chưa có xác nhận lưu thành công từ PAD.")
            if not source.can_run:
                cell.setForeground(QColor("#8A3B32"))
            elif str(source.sqt) in self._skipped_sqt and column != 2:
                cell.setBackground(QColor("#FFF4CC"))
            self.table.setItem(row, column, cell)

    def _clear_remembered(self) -> None:
        if self._clear_saved_callback is None:
            return
        self._clear_saved_callback()
        self._initial_selected_sqt.clear()
        self._set_all(Qt.CheckState.Unchecked)
        if hasattr(self, "restore_label"):
            self.restore_label.setText("Đã xóa lựa chọn ghi nhớ của file và sheet này.")

    def _set_all(self, state: Qt.CheckState) -> None:
        self.table.blockSignals(True)
        try:
            for item in self._checks:
                if item.flags() & Qt.ItemFlag.ItemIsUserCheckable:
                    item.setCheckState(state)
        finally:
            self.table.blockSignals(False)
        self._update_selection_summary()

    def _set_visible(self, state: Qt.CheckState) -> None:
        self.table.blockSignals(True)
        try:
            for row in range(self.table.rowCount()):
                if self.table.isRowHidden(row):
                    continue
                item = self.table.item(row, 0)
                if (
                    item is not None
                    and item.flags() & Qt.ItemFlag.ItemIsUserCheckable
                ):
                    item.setCheckState(state)
        finally:
            self.table.blockSignals(False)
        self._update_selection_summary()

    def _set_filter(self, status: str | None) -> None:
        self._active_filter = status
        self._apply_filter()

    def _apply_filter(self) -> None:
        if not hasattr(self, "table"):
            return
        for row in range(self.table.rowCount()):
            check = self.table.item(row, 0)
            status = check.data(STATUS_ROLE) if check is not None else None
            self.table.setRowHidden(
                row,
                self._active_filter is not None
                and status != self._active_filter,
            )

    def _apply_sort(self, _index: int | None = None) -> None:
        if not hasattr(self, "table"):
            return
        key = str(self.sort_combo.currentData() or "status_asc")
        column, order = self._sort_options[key]
        self.table.sortItems(column, order)
        self._apply_filter()

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if item.column() == 0:
            self._update_selection_summary()

    def _update_selection_summary(self) -> None:
        if not hasattr(self, "selection_summary"):
            return
        selected = [
            item
            for item in self._checks
            if item.flags() & Qt.ItemFlag.ItemIsUserCheckable
            and item.checkState() == Qt.CheckState.Checked
        ]
        imported = sum(
            item.data(STATUS_ROLE) == RPA_STATUS_IMPORTED for item in selected
        )
        not_imported = len(selected) - imported
        if selected:
            self.selection_summary.setText(
                f"Đã chọn: {len(selected)} SQT — "
                f"{not_imported} chưa nhập, {imported} nhập lại"
            )
            if hasattr(self, "run_button"):
                self.run_button.setText(f"Chạy RPA {len(selected)} SQT đã chọn")
        else:
            self.selection_summary.setText("Đã chọn: 0 SQT")
            if hasattr(self, "run_button"):
                self.run_button.setText("Chạy RPA các SQT đã chọn")

    @staticmethod
    def _sqt_sort_key(value: Any) -> tuple[int, int | str]:
        text = str(value).strip()
        return (0, int(text)) if text.isdigit() else (1, text.casefold())

    @property
    def selected_sqt(self) -> list[str]:
        selected: list[str] = []
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if (
                item is not None
                and item.flags() & Qt.ItemFlag.ItemIsUserCheckable
                and item.checkState() == Qt.CheckState.Checked
            ):
                selected.append(str(item.data(Qt.ItemDataRole.UserRole)))
        return selected

    def _validate_and_accept(self) -> None:
        if not self.selected_sqt:
            QMessageBox.warning(
                self,
                "Chưa chọn SQT",
                "Vui lòng chọn ít nhất một SQT hợp lệ để chạy RPA.",
            )
            return
        self.accept()


class RpaLatestDataDialog(AppDialog):
    """Bảng chỉ đọc của đúng payload gần nhất đã gửi sang PAD."""

    COLUMNS = (
        "SQT",
        "Trạng thái trước khi gửi",
        "Dòng BK",
        "Cước MB",
        "N.hạ MB",
        "Cước biển",
        "N.hạ/VS/D/O/Lệnh",
        "Cước MN",
        "Tiền hàng",
        "Công nhân bốc xếp",
        "Lưu cont/Quá tải",
        "Sửa chữa",
        "Tổng",
    )
    AMOUNT_KEYS = (
        "cuoc_bo_dong_hang",
        "nang_ha_dong_hang",
        "cuoc_bien",
        "nang_do_vs_lam_lenh",
        "cuoc_bo_tra_hang",
        "tien_hang",
        "cong_nhan_boc_xep",
        "luu_cont_qua_tai",
        "sua_chua_cont",
    )

    def __init__(
        self,
        payload: Mapping[str, Any],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.payload = dict(payload)
        self.setObjectName("rpaLatestDataDialog")
        self.setWindowTitle("Dữ liệu gần nhất đã gửi sang PAD")
        self.resize(1480, 680)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        sheet_name = str(self.payload.get("sheet_name") or "—")
        run_id = str(self.payload.get("run_id") or "—")
        timestamp = str(
            self.payload.get("launched_at")
            or self.payload.get("created_at")
            or "—"
        )
        title = QLabel(
            f"Sheet {sheet_name} · Run {run_id}\nĐã gửi sang PAD: {timestamp}"
        )
        title.setWordWrap(True)
        title.setStyleSheet("font-size: 12pt; font-weight: 700;")
        layout.addWidget(title)
        note = QLabel(
            "Bảng hiển thị toàn bộ SQT trong payload đã gửi sang PAD gần nhất. "
            "Trạng thái trong bảng là trạng thái tại thời điểm gửi."
        )
        note.setWordWrap(True)
        note.setProperty("muted", True)
        layout.addWidget(note)

        items = self.payload.get("items")
        values = items if isinstance(items, list) else []
        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setObjectName("rpaLatestDataTable")
        self.table.setHorizontalHeaderLabels(list(self.COLUMNS))
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        for source in values:
            if isinstance(source, Mapping):
                self._add_payload_item(source)
        layout.addWidget(self.table, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _add_payload_item(self, source: Mapping[str, Any]) -> None:
        amounts_value = source.get("amounts")
        amounts = amounts_value if isinstance(amounts_value, Mapping) else {}
        amount_values = [amounts.get(key, 0) for key in self.AMOUNT_KEYS]
        source_rows = source.get("source_rows")
        rows_text = (
            ", ".join(str(value) for value in source_rows)
            if isinstance(source_rows, (list, tuple))
            else str(source_rows or "—")
        )
        cells = (
            source.get("sqt"),
            source.get("status_before"),
            rows_text,
            *(_money(value) for value in amount_values),
            _money(sum(int(value or 0) for value in amount_values)),
        )
        row = self.table.rowCount()
        self.table.insertRow(row)
        for column, value in enumerate(cells):
            item = QTableWidgetItem(str(value if value is not None else "—"))
            item.setToolTip(item.text())
            self.table.setItem(row, column, item)


__all__ = ["RpaLatestDataDialog", "RpaSqtSelectionDialog"]
