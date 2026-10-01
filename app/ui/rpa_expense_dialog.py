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
    QLineEdit,
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
from .presentation import dialog_intro, fit_window_to_screen


STATUS_ROLE = Qt.ItemDataRole.UserRole + 1
PENDING_ROLE = Qt.ItemDataRole.UserRole + 2
LATEST_PAD_ROLE = Qt.ItemDataRole.UserRole + 3


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
        "Số quyết toán",
        "Trạng thái nhập",
        "Nhóm",
        "Dòng BK",
        "Cước MB",
        "N.hạ MB",
        "Cước biển",
        "N.hạ/VS/D/O/Lệnh",
        "Cước MN",
        "Lưu cont/Quá tải",
        "Sửa chữa",
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
        self._pending_sqt = {
            str(value) for value in getattr(plan, "pending_sqt", ())
        }
        self._latest_pad_sqt = {
            str(value) for value in getattr(plan, "latest_pad_sqt", ())
        }
        available_sqt = {str(item.sqt) for item in self.plan.items}
        runnable_sqt = {
            str(item.sqt) for item in self.plan.items if bool(item.can_run)
        }
        self._pending_sqt.intersection_update(available_sqt)
        self._latest_pad_sqt.intersection_update(available_sqt)
        self._selectable_pending_sqt = self._pending_sqt & runnable_sqt
        self._selectable_latest_pad_sqt = self._latest_pad_sqt & runnable_sqt
        self._initial_selected_sqt.update(self._selectable_pending_sqt)
        self._initial_selected_sqt.update(self._selectable_latest_pad_sqt)
        self.restore_info = dict(restore_info or {})
        self._skipped_sqt = {
            str(value) for value in self.restore_info.get("skipped_sqt", ())
        }
        self._clear_saved_callback = clear_saved_callback
        self._checks: list[QTableWidgetItem] = []
        self._active_group_filter = (
            "pending"
            if self._pending_sqt
            else "latest"
            if self._latest_pad_sqt
            else "all"
        )
        self._active_status_filter: str | None = None
        self._sort_options = {
            "status_asc": (2, Qt.SortOrder.AscendingOrder),
            "status_desc": (2, Qt.SortOrder.DescendingOrder),
            "sqt_asc": (1, Qt.SortOrder.AscendingOrder),
            "sqt_desc": (1, Qt.SortOrder.DescendingOrder),
            "rows_asc": (4, Qt.SortOrder.AscendingOrder),
            "rows_desc": (4, Qt.SortOrder.DescendingOrder),
        }
        self.setObjectName("rpaSqtSelectionDialog")
        self.setWindowTitle("Chọn số quyết toán để nhập")
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
                background: #EEF5F7;
                border: 1px solid #D3E4E8;
                border-radius: 8px;
                padding: 9px 12px;
            }
            """
        )
        self._build_ui()
        fit_window_to_screen(self, 1200, 650)

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        intro, _ = dialog_intro(
            f"Chọn số quyết toán trong trang tính {self.plan.sheet_name}",
            "Có thể chọn lại số đã nhập nếu cần chạy lại. "
            "Trạng thái chỉ đổi thành “Đã nhập” khi phần mềm quyết toán xác nhận lưu thành công.",
        )
        layout.addWidget(intro)

        if self._pending_sqt or self._latest_pad_sqt:
            overlap = len(
                self._selectable_pending_sqt & self._selectable_latest_pad_sqt
            )
            selected_count = len(
                self._selectable_pending_sqt | self._selectable_latest_pad_sqt
            )
            self.group_summary_label = QLabel(
                f"Đã chọn sẵn {selected_count} số quyết toán: "
                f"{len(self._selectable_pending_sqt)} chờ nhập + "
                f"{len(self._selectable_latest_pad_sqt)} thuộc lượt nhập gần nhất"
                + (f" − {overlap} trùng nhau." if overlap else ".")
            )
            self.group_summary_label.setWordWrap(True)
            self.group_summary_label.setProperty("status", "info")
            layout.addWidget(self.group_summary_label)

        if self.restore_info.get("found") and not (
            self._pending_sqt or self._latest_pad_sqt
        ):
            restored = int(self.restore_info.get("restored_count", 0) or 0)
            saved = int(self.restore_info.get("saved_count", 0) or 0)
            skipped = max(0, saved - restored)
            restore_row = QHBoxLayout()
            message = f"Đã khôi phục {restored}/{saved} số quyết toán từ lần chạy gần nhất."
            if skipped:
                message += f" Có {skipped} số đã thay đổi hoặc không còn hợp lệ."
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

        group_filters = QHBoxLayout()
        group_filters.setSpacing(7)
        group_filters.addWidget(QLabel("Nhóm hiển thị:"))
        self.group_filter_group = QButtonGroup(self)
        self.group_filter_group.setExclusive(True)
        group_specs = (
            ("Tất cả", "all", len(self.plan.items), "allRpaGroupFilterButton"),
            (
                "Chờ nhập",
                "pending",
                len(self._pending_sqt),
                "pendingRpaGroupFilterButton",
            ),
            (
                "Lượt nhập gần nhất",
                "latest",
                len(self._latest_pad_sqt),
                "latestRpaGroupFilterButton",
            ),
        )
        self.group_filter_buttons: dict[str, QPushButton] = {}
        for label, group_name, count, object_name in group_specs:
            button = QPushButton(f"{label} ({count})")
            button.setObjectName(object_name)
            button.setProperty("filterChip", True)
            button.setCheckable(True)
            button.clicked.connect(
                lambda checked, value=group_name: checked
                and self._set_group_filter(value)
            )
            self.group_filter_group.addButton(button)
            self.group_filter_buttons[group_name] = button
            group_filters.addWidget(button)
        self.group_filter_buttons[self._active_group_filter].setChecked(True)
        group_filters.addStretch()
        layout.addLayout(group_filters)

        table_tools = QHBoxLayout()
        table_tools.setSpacing(7)
        table_tools.addWidget(QLabel("Tìm số quyết toán:"))
        self.sqt_search = QLineEdit()
        self.sqt_search.setObjectName("rpaSqtSearchInput")
        self.sqt_search.setPlaceholderText("Nhập số QT cần tìm...")
        self.sqt_search.setClearButtonEnabled(True)
        self.sqt_search.setMaximumWidth(280)
        self.sqt_search.textChanged.connect(self._apply_filter)
        table_tools.addWidget(self.sqt_search)
        table_tools.addStretch()
        table_tools.addWidget(QLabel("Sắp xếp:"))
        self.sort_combo = QComboBox()
        self.sort_combo.setObjectName("rpaSortCombo")
        self.sort_combo.addItem("Trạng thái: Chưa nhập trước", "status_asc")
        self.sort_combo.addItem("Trạng thái: Đã nhập trước", "status_desc")
        self.sort_combo.addItem("Số quyết toán tăng dần", "sqt_asc")
        self.sort_combo.addItem("Số quyết toán giảm dần", "sqt_desc")
        self.sort_combo.addItem("Dòng BK tăng dần", "rows_asc")
        self.sort_combo.addItem("Dòng BK giảm dần", "rows_desc")
        self.sort_combo.currentIndexChanged.connect(self._apply_sort)
        table_tools.addWidget(self.sort_combo)

        filters = QHBoxLayout()
        filters.setSpacing(7)
        filters.addWidget(QLabel("Trạng thái nhập:"))
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
                lambda checked, value=status: checked
                and self._set_status_filter(value)
            )
            self.filter_group.addButton(button)
            self.filter_buttons[status] = button
            filters.addWidget(button)
        self.filter_buttons[None].setChecked(True)
        filters.addStretch()
        layout.addLayout(filters)
        layout.addLayout(table_tools)

        selection = QHBoxLayout()
        select_all = QPushButton("Chọn tất cả")
        clear_all = QPushButton("Bỏ chọn tất cả")
        select_all.setObjectName("selectAllRpaSqtButton")
        clear_all.setObjectName("clearAllRpaSqtButton")
        select_all.clicked.connect(lambda: self._set_all(Qt.CheckState.Checked))
        clear_all.clicked.connect(
            lambda: self._set_all(Qt.CheckState.Unchecked)
        )
        select_all.setToolTip("Chọn toàn bộ số quyết toán hợp lệ, kể cả các dòng đang bị lọc")
        clear_all.setToolTip("Bỏ chọn toàn bộ số quyết toán, kể cả các dòng đang bị lọc")
        selection.addWidget(select_all)
        selection.addWidget(clear_all)
        selection.addStretch()
        self.selection_summary = QLabel()
        self.selection_summary.setObjectName("rpaSelectionSummaryLabel")
        layout.addLayout(selection)
        self.selection_summary.setWordWrap(True)
        layout.addWidget(self.selection_summary)

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
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.setHorizontalScrollMode(
            QAbstractItemView.ScrollMode.ScrollPerPixel
        )
        self.table.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOn
        )
        self.table.setWordWrap(False)
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
        self.run_button.setText("Bắt đầu nhập các số đã chọn")
        self.run_button.setProperty("primary", True)
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
        is_pending = str(source.sqt) in self._pending_sqt
        is_latest = str(source.sqt) in self._latest_pad_sqt
        check.setData(PENDING_ROLE, is_pending)
        check.setData(LATEST_PAD_ROLE, is_latest)
        check.setToolTip(
            "Phần mềm quyết toán đã xác nhận lưu thành công; vẫn có thể chọn để nhập lại."
            if source.status == RPA_STATUS_IMPORTED
            else "Chưa có xác nhận lưu thành công từ phần mềm quyết toán."
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
        group_display = (
            "Cả hai"
            if is_pending and is_latest
            else "Chờ nhập"
            if is_pending
            else "Lượt nhập gần nhất"
            if is_latest
            else ""
        )
        sqt_sort = self._sqt_sort_key(source.sqt)
        group_sort = (
            0
            if is_pending and is_latest
            else 1
            if is_pending
            else 2
            if is_latest
            else 3
        )
        source_rows = tuple(int(value) for value in source.source_rows)
        values = (
            source.sqt,
            status_display,
            group_display,
            ", ".join(str(value) for value in source.source_rows),
            _money(amounts.cuoc_bo_dong_hang),
            _money(amounts.nang_ha_dong_hang),
            _money(amounts.cuoc_bien),
            _money(amounts.nang_do_vs_lam_lenh),
            _money(amounts.cuoc_bo_tra_hang),
            _money(amounts.luu_cont_qua_tai),
            _money(amounts.sua_chua_cont),
            source.validation_message or "Hợp lệ",
        )
        sort_values = (
            sqt_sort,
            (
                1 if source.status == RPA_STATUS_IMPORTED else 0,
                sqt_sort,
            ),
            (group_sort, sqt_sort),
            source_rows,
            amounts.cuoc_bo_dong_hang,
            amounts.nang_ha_dong_hang,
            amounts.cuoc_bien,
            amounts.nang_do_vs_lam_lenh,
            amounts.cuoc_bo_tra_hang,
            amounts.luu_cont_qua_tai,
            amounts.sua_chua_cont,
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
                        "Phần mềm quyết toán đã xác nhận lưu thành công; "
                        "vẫn có thể chọn để nhập lại."
                    )
                else:
                    cell.setForeground(QColor("#1D4ED8"))
                    cell.setBackground(QColor("#EFF6FF"))
                    cell.setToolTip("Chưa có xác nhận lưu thành công từ phần mềm quyết toán.")
            elif column == 3 and group_display:
                font = cell.font()
                font.setBold(True)
                cell.setFont(font)
                cell.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if is_pending and is_latest:
                    cell.setForeground(QColor("#15803D"))
                    cell.setBackground(QColor("#DCFCE7"))
                elif is_pending:
                    cell.setForeground(QColor("#1D4ED8"))
                    cell.setBackground(QColor("#DBEAFE"))
                else:
                    cell.setForeground(QColor("#6D28D9"))
                    cell.setBackground(QColor("#EDE9FE"))
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
            self.restore_label.setText("Đã xóa lựa chọn ghi nhớ của trang tính này.")

    def _set_all(self, state: Qt.CheckState) -> None:
        self.table.blockSignals(True)
        try:
            for item in self._checks:
                if item.flags() & Qt.ItemFlag.ItemIsUserCheckable:
                    item.setCheckState(state)
        finally:
            self.table.blockSignals(False)
        self._update_selection_summary()

    def _set_group_filter(self, group_name: str) -> None:
        self._active_group_filter = group_name
        self._apply_filter()

    def _set_status_filter(self, status: str | None) -> None:
        self._active_status_filter = status
        self._apply_filter()

    # Alias giữ tương thích với adapter/test cũ.
    def _set_filter(self, status: str | None) -> None:
        self._set_status_filter(status)

    def _apply_filter(self) -> None:
        if not hasattr(self, "table"):
            return
        search_text = (
            self.sqt_search.text().strip().casefold()
            if hasattr(self, "sqt_search")
            else ""
        )
        for row in range(self.table.rowCount()):
            check = self.table.item(row, 0)
            status = check.data(STATUS_ROLE) if check is not None else None
            pending = bool(check.data(PENDING_ROLE)) if check is not None else False
            latest = bool(check.data(LATEST_PAD_ROLE)) if check is not None else False
            group_matches = (
                self._active_group_filter == "all"
                or (self._active_group_filter == "pending" and pending)
                or (self._active_group_filter == "latest" and latest)
            )
            status_matches = (
                self._active_status_filter is None
                or status == self._active_status_filter
            )
            sqt_item = self.table.item(row, 1)
            search_matches = (
                not search_text
                or (
                    sqt_item is not None
                    and search_text in sqt_item.text().strip().casefold()
                )
            )
            self.table.setRowHidden(
                row,
                not (group_matches and status_matches and search_matches),
            )
        self._update_selection_summary()

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
        visible_selected = sum(
            not self.table.isRowHidden(item.row()) for item in selected
        )
        if selected:
            selected_sqt = {
                str(item.data(Qt.ItemDataRole.UserRole)) for item in selected
            }
            total_amount = sum(
                item.amounts.total
                for item in self.plan.items
                if str(item.sqt) in selected_sqt
            )
            self.selection_summary.setText(
                f"Đã chọn: {len(selected)} số quyết toán — "
                f"đang hiển thị {visible_selected}; "
                f"{not_imported} chưa nhập, {imported} nhập lại · "
                f"Tổng các khoản chi: {_money(total_amount)} đ"
            )
            if hasattr(self, "run_button"):
                self.run_button.setText(f"Bắt đầu nhập {len(selected)} số đã chọn")
        else:
            self.selection_summary.setText("Chưa chọn số quyết toán nào")
            if hasattr(self, "run_button"):
                self.run_button.setText("Bắt đầu nhập các số đã chọn")

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
                "Chưa chọn số quyết toán",
                "Vui lòng chọn ít nhất một số quyết toán hợp lệ để bắt đầu nhập.",
            )
            return
        self.accept()


class RpaLatestDataDialog(AppDialog):
    """Bảng chỉ đọc của đúng payload gần nhất đã gửi sang PAD."""

    COLUMNS = (
        "Số quyết toán",
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
        self.setWindowTitle("Dữ liệu của lần nhập gần nhất")
        self._build_ui()
        fit_window_to_screen(self, 1200, 650)

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        sheet_name = str(self.payload.get("sheet_name") or "—")
        run_id = str(self.payload.get("run_id") or "—")
        timestamp = str(
            self.payload.get("launched_at")
            or self.payload.get("created_at")
            or "—"
        )
        intro, _ = dialog_intro(
            f"Trang tính {sheet_name} · Gửi lúc {timestamp}",
            "Đây là các số quyết toán đã chuyển sang công cụ nhập trong lần gần nhất. "
            "Trạng thái bên dưới là trạng thái tại thời điểm gửi."
        )
        intro.setToolTip(f"Mã lượt xử lý: {run_id}")
        layout.addWidget(intro)

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
        buttons.button(QDialogButtonBox.StandardButton.Close).setText("Đóng")
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
