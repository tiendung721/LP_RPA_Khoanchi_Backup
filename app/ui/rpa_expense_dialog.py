"""Hộp thoại xem và chọn SQT trước khi gọi PAD."""

from __future__ import annotations

from collections import OrderedDict
from collections.abc import Mapping
from typing import Any, Callable, Iterable

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.rpa_expense.contracts import (
    RPA_FEE_GROUPS,
    RPA_STATUS_IMPORTED,
    RPA_STATUS_NOT_IMPORTED,
)
from app.services.excel.resolvers import MonthSheetService

from .app_dialog import AppDialog
from .presentation import dialog_intro, fit_window_to_screen


STATUS_ROLE = Qt.ItemDataRole.UserRole + 1
LATEST_BK_ROLE = Qt.ItemDataRole.UserRole + 2
LATEST_PAD_ROLE = Qt.ItemDataRole.UserRole + 3
SOURCE_ROLE = Qt.ItemDataRole.UserRole + 4


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
    MULTI_COLUMNS = (
        "Chọn",
        "Số quyết toán",
        "Sheet BK",
        "HĐ theo loại phí",
        "Trạng thái nhập",
        "Tổng khoản chi",
    )
    COLUMNS = (
        "Chọn",
        "Số quyết toán",
        "Số hóa đơn",
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
        self._multiple_sheets = not bool(plan.sheet_name)
        self._initial_selected_sqt = {
            str(value) for value in initial_selected_sqt
        }
        self._latest_bk_sqt = {
            str(value) for value in getattr(plan, "latest_bk_sqt", ())
        }
        self._latest_pad_sqt = {
            str(value) for value in getattr(plan, "latest_pad_sqt", ())
        }
        available_sqt = {str(item.sqt) for item in self.plan.items}
        runnable_sqt = {
            str(item.sqt) for item in self.plan.items if bool(item.can_run)
        }
        self._latest_bk_sqt.intersection_update(available_sqt)
        self._latest_pad_sqt.intersection_update(available_sqt)
        self._selectable_latest_bk_sqt = self._latest_bk_sqt & runnable_sqt
        self._selectable_latest_pad_sqt = self._latest_pad_sqt & runnable_sqt
        self._initial_selected_sqt.update(self._selectable_latest_bk_sqt)
        self.restore_info = dict(restore_info or {})
        self._skipped_sqt = {
            str(value) for value in self.restore_info.get("skipped_sqt", ())
        }
        self._clear_saved_callback = clear_saved_callback
        self._checks: list[QTableWidgetItem] = []
        self._active_group_filter = (
            "bk" if self._latest_bk_sqt and not self._multiple_sheets else "all"
        )
        self._active_status_filter: str | None = None
        self._sort_options = {
            "status_asc": (3, Qt.SortOrder.AscendingOrder),
            "status_desc": (3, Qt.SortOrder.DescendingOrder),
            "sqt_asc": (1, Qt.SortOrder.AscendingOrder),
            "sqt_desc": (1, Qt.SortOrder.DescendingOrder),
            "rows_asc": (5, Qt.SortOrder.AscendingOrder),
            "rows_desc": (5, Qt.SortOrder.DescendingOrder),
        }
        if self._multiple_sheets:
            self._sort_options = {
                "status_asc": (4, Qt.SortOrder.AscendingOrder),
                "status_desc": (4, Qt.SortOrder.DescendingOrder),
                "sqt_asc": (1, Qt.SortOrder.AscendingOrder),
                "sqt_desc": (1, Qt.SortOrder.DescendingOrder),
                "sheet_asc": (2, Qt.SortOrder.AscendingOrder),
                "sheet_desc": (2, Qt.SortOrder.DescendingOrder),
                "total_desc": (5, Qt.SortOrder.DescendingOrder),
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
        fit_window_to_screen(
            self, 1360 if self._multiple_sheets else 1200,
            760 if self._multiple_sheets else 650,
        )
        if self._multiple_sheets:
            self._update_responsive_layout()

    def _build_ui(self) -> None:
        if self._multiple_sheets:
            self._build_multi_sheet_ui()
            return
        layout = QVBoxLayout(self)
        intro, _ = dialog_intro(
            (
                "Chọn số quyết toán từ các sheet tháng trong BK"
                if self._multiple_sheets
                else f"Chọn số quyết toán trong trang tính {self.plan.sheet_name}"
            ),
            "Có thể chọn lại số đã nhập nếu cần chạy lại. "
            "Trạng thái chỉ đổi thành “Đã nhập” khi phần mềm quyết toán xác nhận lưu thành công.",
        )
        layout.addWidget(intro)

        if self._latest_bk_sqt or self._latest_pad_sqt:
            self.group_summary_label = QLabel(
                f"Vừa ghi BK: {len(self._latest_bk_sqt)} số quyết toán "
                f"(chọn sẵn {len(self._selectable_latest_bk_sqt)} hợp lệ). "
                f"Lượt nhập RPA gần nhất: {len(self._latest_pad_sqt)} số quyết toán."
            )
            self.group_summary_label.setWordWrap(True)
            self.group_summary_label.setProperty("status", "info")
            layout.addWidget(self.group_summary_label)

        if self.restore_info.get("found") and not (
            self._latest_bk_sqt or self._latest_pad_sqt
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
                "Vừa ghi BK",
                "bk",
                len(self._latest_bk_sqt),
                "latestBkRpaGroupFilterButton",
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
        table_tools.addWidget(QLabel("Tìm SQT hoặc số hóa đơn:"))
        self.sqt_search = QLineEdit()
        self.sqt_search.setObjectName("rpaSqtSearchInput")
        self.sqt_search.setPlaceholderText("Nhập SQT hoặc số hóa đơn...")
        self.sqt_search.setClearButtonEnabled(True)
        self.sqt_search.setMaximumWidth(280)
        self.sqt_search.textChanged.connect(self._apply_filter)
        table_tools.addWidget(self.sqt_search)
        table_tools.addStretch()
        if self._multiple_sheets:
            table_tools.addWidget(QLabel("Sheet BK:"))
            self.sheet_filter = QComboBox()
            self.sheet_filter.setObjectName("rpaSheetFilterCombo")
            self.sheet_filter.setMinimumWidth(155)
            counts: dict[str, int] = {}
            for item in self.plan.items:
                counts[item.sheet_name] = counts.get(item.sheet_name, 0) + 1
            self.sheet_filter.addItem(f"Tất cả ({len(self.plan.items)})", None)
            months = MonthSheetService()
            for sheet_name in sorted(
                counts,
                key=lambda value: (
                    (parsed[1], parsed[0])
                    if (parsed := months.parse_target_sheet(value)) else (0, 0)
                ),
                reverse=True,
            ):
                self.sheet_filter.addItem(
                    f"{sheet_name} ({counts[sheet_name]})", sheet_name
                )
            self.sheet_filter.currentIndexChanged.connect(self._apply_filter)
            table_tools.addWidget(self.sheet_filter)
        table_tools.addWidget(QLabel("Sắp xếp:"))
        self.sort_combo = QComboBox()
        self.sort_combo.setObjectName("rpaSortCombo")
        self.sort_combo.setMinimumWidth(210)
        self.sort_combo.addItem("Trạng thái: Chưa nhập trước", "status_asc")
        self.sort_combo.addItem("Trạng thái: Đã nhập trước", "status_desc")
        self.sort_combo.addItem("Số quyết toán tăng dần", "sqt_asc")
        self.sort_combo.addItem("Số quyết toán giảm dần", "sqt_desc")
        if self._multiple_sheets:
            self.sort_combo.addItem("Sheet BK tăng dần", "sheet_asc")
            self.sort_combo.addItem("Sheet BK giảm dần", "sheet_desc")
            self.sort_combo.addItem("Tổng tiền giảm dần", "total_desc")
        else:
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
        select_all.setToolTip("Chỉ chọn các số quyết toán hợp lệ đang hiển thị")
        clear_all.setToolTip("Chỉ bỏ chọn các số quyết toán đang hiển thị")
        selection.addWidget(select_all)
        selection.addWidget(clear_all)
        selection.addStretch()
        self.selection_summary = QLabel()
        self.selection_summary.setObjectName("rpaSelectionSummaryLabel")
        layout.addLayout(selection)
        self.selection_summary.setWordWrap(True)
        layout.addWidget(self.selection_summary)

        columns = self.MULTI_COLUMNS if self._multiple_sheets else self.COLUMNS
        self.table = QTableWidget(0, len(columns))
        self.table.setObjectName("rpaSqtTable")
        self.table.setHorizontalHeaderLabels(list(columns))
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        if self._multiple_sheets:
            self.table.horizontalHeader().setSectionResizeMode(
                3, QHeaderView.ResizeMode.Stretch
            )
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.setHorizontalScrollMode(
            QAbstractItemView.ScrollMode.ScrollPerPixel
        )
        self.table.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOn
        )
        self.table.setWordWrap(self._multiple_sheets)
        layout.addWidget(self.table, 1)

        for item in self.plan.items:
            self._add_item(item)

        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().sortIndicatorChanged.connect(
            lambda *_: QTimer.singleShot(0, self._apply_filter)
        )
        self.table.itemChanged.connect(self._on_item_changed)
        if self._multiple_sheets:
            self.table.currentCellChanged.connect(self._show_detail)
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

    def _build_multi_sheet_ui(self) -> None:
        """Danh sách SQT luôn chiếm phần lớn chiều cao cửa sổ."""
        self.setStyleSheet(self.styleSheet() + """
            QLabel#rpaCompactTitle { font-size: 18px; font-weight: 700; color: #12304E; }
            QLabel#rpaCountBadge {
                color: #1D4ED8; background: #E8F1FF; border-radius: 10px;
                padding: 4px 9px; font-weight: 600;
            }
            QWidget#rpaDetailPanel { background: #F8FBFF; border: 1px solid #D6E2EF; }
            QFrame#rpaFeeCard {
                background: #FFFFFF; border: 1px solid #D8E4F2;
                border-radius: 6px;
            }
            QLabel#rpaSelectionSummaryLabel {
                background: #F1F6FC; border: 1px solid #D6E2EF;
                padding: 7px 10px; border-radius: 6px;
            }
            QPushButton#rpaDetailToggle { color: #1D4ED8; }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 9, 12, 9)
        layout.setSpacing(7)

        heading = QHBoxLayout()
        heading.setSpacing(9)
        title = QLabel("Chọn số quyết toán")
        title.setObjectName("rpaCompactTitle")
        heading.addWidget(title)
        count = QLabel(f"{len(self.plan.items)} SQT trong BK")
        count.setObjectName("rpaCountBadge")
        count.setToolTip(
            f"Vừa ghi BK: {len(self._latest_bk_sqt)} SQT; "
            f"lượt nhập gần nhất: {len(self._latest_pad_sqt)} SQT."
        )
        heading.addWidget(count)
        help_button = QPushButton("i")
        help_button.setObjectName("rpaSelectionHelpButton")
        help_button.setFixedSize(23, 23)
        help_button.setToolTip(
            "Có thể chọn lại SQT đã nhập để chạy lại. Trạng thái Đã nhập "
            "chỉ cập nhật khi phần mềm quyết toán xác nhận lưu thành công."
        )
        heading.addWidget(help_button)
        heading.addStretch()
        layout.addLayout(heading)

        if self.restore_info.get("found") and not (
            self._latest_bk_sqt or self._latest_pad_sqt
        ):
            restored = int(self.restore_info.get("restored_count", 0) or 0)
            saved = int(self.restore_info.get("saved_count", 0) or 0)
            restore_row = QHBoxLayout()
            self.restore_label = QLabel(
                f"Đã khôi phục {restored}/{saved} SQT từ lần chạy gần nhất."
            )
            self.restore_label.setToolTip(
                "Các SQT không còn hợp lệ hoặc đã thay đổi sẽ được bỏ qua."
            )
            restore_row.addWidget(self.restore_label, 1)
            if self._clear_saved_callback is not None:
                clear_saved = QPushButton("Xóa lựa chọn đã nhớ")
                clear_saved.setObjectName("clearRememberedRpaSelectionButton")
                clear_saved.clicked.connect(self._clear_remembered)
                restore_row.addWidget(clear_saved)
            layout.addLayout(restore_row)

        self.sqt_search = QLineEdit()
        self.sqt_search.setObjectName("rpaSqtSearchInput")
        self.sqt_search.setPlaceholderText("Nhập SQT hoặc số hóa đơn...")
        self.sqt_search.setClearButtonEnabled(True)
        self.sqt_search.textChanged.connect(self._apply_filter)

        self.group_filter = QComboBox()
        self.group_filter.setObjectName("rpaGroupFilterCombo")
        self.group_filter.addItem(f"Tất cả ({len(self.plan.items)})", "all")
        self.group_filter.addItem(f"Vừa ghi BK ({len(self._latest_bk_sqt)})", "bk")
        self.group_filter.addItem(
            f"Lượt nhập gần nhất ({len(self._latest_pad_sqt)})", "latest"
        )
        self.group_filter.setCurrentIndex(
            self.group_filter.findData(self._active_group_filter)
        )
        self.group_filter.currentIndexChanged.connect(
            lambda *_: self._set_group_filter(str(self.group_filter.currentData()))
        )

        self.status_filter = QComboBox()
        self.status_filter.setObjectName("rpaStatusFilterCombo")
        imported_count = sum(
            item.status == RPA_STATUS_IMPORTED for item in self.plan.items
        )
        self.status_filter.addItem(f"Tất cả ({len(self.plan.items)})", None)
        self.status_filter.addItem(
            f"Chưa nhập ({len(self.plan.items) - imported_count})",
            RPA_STATUS_NOT_IMPORTED,
        )
        self.status_filter.addItem(
            f"Đã nhập ({imported_count})", RPA_STATUS_IMPORTED
        )
        self.status_filter.currentIndexChanged.connect(
            lambda *_: self._set_status_filter(self.status_filter.currentData())
        )

        self.sheet_filter = QComboBox()
        self.sheet_filter.setObjectName("rpaSheetFilterCombo")
        counts: dict[str, int] = {}
        for item in self.plan.items:
            counts[item.sheet_name] = counts.get(item.sheet_name, 0) + 1
        self.sheet_filter.addItem(f"Tất cả ({len(self.plan.items)})", None)
        months = MonthSheetService()
        for sheet_name in sorted(
            counts,
            key=lambda value: (
                (parsed[1], parsed[0])
                if (parsed := months.parse_target_sheet(value)) else (0, 0)
            ),
            reverse=True,
        ):
            self.sheet_filter.addItem(
                f"{sheet_name} ({counts[sheet_name]})", sheet_name
            )
        self.sheet_filter.currentIndexChanged.connect(self._apply_filter)

        self.sort_combo = QComboBox()
        self.sort_combo.setObjectName("rpaSortCombo")
        for label, key in (
            ("Trạng thái: Chưa nhập trước", "status_asc"),
            ("Trạng thái: Đã nhập trước", "status_desc"),
            ("Số quyết toán tăng dần", "sqt_asc"),
            ("Số quyết toán giảm dần", "sqt_desc"),
            ("Sheet BK tăng dần", "sheet_asc"),
            ("Sheet BK giảm dần", "sheet_desc"),
            ("Tổng tiền giảm dần", "total_desc"),
        ):
            self.sort_combo.addItem(label, key)
        self.sort_combo.currentIndexChanged.connect(self._apply_sort)

        def field(label: str, control: QWidget) -> QWidget:
            wrapper = QWidget()
            wrapper_layout = QVBoxLayout(wrapper)
            wrapper_layout.setContentsMargins(0, 0, 0, 0)
            wrapper_layout.setSpacing(3)
            caption = QLabel(label)
            caption.setStyleSheet("color: #59708C; font-size: 11px; font-weight: 600;")
            wrapper_layout.addWidget(caption)
            wrapper_layout.addWidget(control)
            control.setMinimumHeight(32)
            return wrapper

        self._toolbar_fields = (
            field("Tìm SQT hoặc số HĐ", self.sqt_search),
            field("Nhóm hiển thị", self.group_filter),
            field("Trạng thái nhập", self.status_filter),
            field("Sheet BK", self.sheet_filter),
            field("Sắp xếp", self.sort_combo),
        )
        bulk = QWidget()
        bulk_layout = QHBoxLayout(bulk)
        bulk_layout.setContentsMargins(0, 0, 0, 0)
        bulk_layout.setSpacing(5)
        select_all = QPushButton("Chọn hiển thị")
        clear_all = QPushButton("Bỏ chọn")
        select_all.setObjectName("selectAllRpaSqtButton")
        clear_all.setObjectName("clearAllRpaSqtButton")
        select_all.setToolTip("Chỉ chọn SQT hợp lệ đang hiển thị")
        clear_all.setToolTip("Chỉ bỏ chọn SQT đang hiển thị")
        select_all.clicked.connect(lambda: self._set_all(Qt.CheckState.Checked))
        clear_all.clicked.connect(lambda: self._set_all(Qt.CheckState.Unchecked))
        bulk_layout.addWidget(select_all)
        bulk_layout.addWidget(clear_all)
        self._toolbar_bulk = bulk
        self._toolbar_grid = QGridLayout()
        self._toolbar_grid.setContentsMargins(0, 0, 0, 0)
        self._toolbar_grid.setHorizontalSpacing(8)
        self._toolbar_grid.setVerticalSpacing(6)
        layout.addLayout(self._toolbar_grid)

        selection_row = QHBoxLayout()
        self.selection_summary = QLabel()
        self.selection_summary.setObjectName("rpaSelectionSummaryLabel")
        selection_row.addWidget(self.selection_summary, 1)
        self.detail_toggle = QPushButton("Ẩn chi tiết số HĐ")
        self.detail_toggle.setObjectName("rpaDetailToggle")
        self.detail_toggle.clicked.connect(self._toggle_detail)
        selection_row.addWidget(self.detail_toggle)
        layout.addLayout(selection_row)

        self.table = QTableWidget(0, len(self.MULTI_COLUMNS))
        self.table.setObjectName("rpaSqtTable")
        self.table.setHorizontalHeaderLabels(list(self.MULTI_COLUMNS))
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        for column, width in enumerate((50, 106, 92, 0, 132, 136)):
            if column == 3:
                header.setSectionResizeMode(column, QHeaderView.ResizeMode.Stretch)
            else:
                header.setSectionResizeMode(column, QHeaderView.ResizeMode.Fixed)
                header.resizeSection(column, width)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table.setWordWrap(True)

        self.detail_panel = QWidget()
        self.detail_panel.setObjectName("rpaDetailPanel")
        self.detail_panel.setMinimumWidth(340)
        detail_layout = QVBoxLayout(self.detail_panel)
        detail_layout.setContentsMargins(10, 10, 10, 10)
        detail_layout.setSpacing(8)
        self.detail_title = QLabel("Chọn một dòng SQT để xem phí và HĐ.")
        self.detail_title.setObjectName("rpaFeeDetailTitle")
        self.detail_title.setWordWrap(True)
        detail_layout.addWidget(self.detail_title)
        caption = QLabel("SỐ HĐ ĐI CÙNG TỪNG LOẠI PHÍ")
        caption.setStyleSheet("color: #66809D; font-size: 11px; font-weight: 700;")
        detail_layout.addWidget(caption)
        self.detail_table = QTableWidget(0, 4, self.detail_panel)
        self.detail_table.setObjectName("rpaFeeDetailTable")
        self.detail_table.setHorizontalHeaderLabels(
            ["Loại phí", "Số tiền", "Số HĐ", "Dòng BK"]
        )
        self.detail_table.setEditTriggers(
            QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.detail_table.setSelectionMode(
            QAbstractItemView.SelectionMode.NoSelection
        )
        self.detail_table.verticalHeader().setVisible(False)
        self.detail_table.setWordWrap(True)
        detail_header = self.detail_table.horizontalHeader()
        detail_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        for column, width in enumerate((120, 95, 95, 62)):
            detail_header.resizeSection(column, width)
        self.detail_table.hide()
        self.detail_cards = QScrollArea()
        self.detail_cards.setWidgetResizable(True)
        self.detail_cards.setFrameShape(QFrame.Shape.NoFrame)
        cards_container = QWidget()
        self.detail_cards_layout = QVBoxLayout(cards_container)
        self.detail_cards_layout.setContentsMargins(0, 0, 0, 0)
        self.detail_cards_layout.setSpacing(7)
        self.detail_cards_layout.addStretch()
        self.detail_cards.setWidget(cards_container)
        detail_layout.addWidget(self.detail_cards, 1)

        self.body_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.body_splitter.setChildrenCollapsible(False)
        self.body_splitter.addWidget(self.table)
        self.body_splitter.addWidget(self.detail_panel)
        self.body_splitter.setSizes([930, 390])
        layout.addWidget(self.body_splitter, 1)

        for item in self.plan.items:
            self._add_item(item)
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().sortIndicatorChanged.connect(
            lambda *_: QTimer.singleShot(0, self._apply_filter)
        )
        self.table.itemChanged.connect(self._on_item_changed)
        self.table.currentCellChanged.connect(self._show_detail)

        footer = QHBoxLayout()
        self.footer_summary = QLabel()
        self.footer_summary.setStyleSheet("color: #315779; font-weight: 600;")
        footer.addWidget(self.footer_summary, 1)
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
        footer.addWidget(buttons)
        layout.addLayout(footer)

        self._detail_user_hidden = False
        self._toolbar_mode = ""
        self._detail_popup: QDialog | None = None
        self._arrange_multi_toolbar("compact")
        self._apply_sort()
        self._update_selection_summary()

    def _arrange_multi_toolbar(self, mode: str) -> None:
        if self._toolbar_mode == mode:
            return
        self._toolbar_mode = mode
        for widget in (*self._toolbar_fields, self._toolbar_bulk):
            self._toolbar_grid.removeWidget(widget)
        positions = (
            [(0, column) for column in range(6)]
            if mode == "wide" else
            [(0, 0), (0, 1), (0, 2), (1, 0), (1, 1), (1, 2)]
        )
        for widget, (row, column) in zip(
            (*self._toolbar_fields, self._toolbar_bulk), positions
        ):
            self._toolbar_grid.addWidget(widget, row, column)
        for column in range(6):
            self._toolbar_grid.setColumnStretch(
                column, (2 if column == 0 else 1)
                if mode == "wide" or column < 3 else 0
            )

    def _update_responsive_layout(self) -> None:
        if not self._multiple_sheets or not hasattr(self, "detail_panel"):
            return
        narrow = self.width() < 1180
        self._arrange_multi_toolbar("compact" if self.width() < 1390 else "wide")
        show_panel = not narrow and not self._detail_user_hidden
        self.detail_panel.setVisible(show_panel)
        self.detail_toggle.setText(
            "Ẩn chi tiết số HĐ" if show_panel else "Xem chi tiết số HĐ"
        )
        if not narrow and self._detail_popup is not None:
            self._detail_popup.hide()

    def _refresh_detail_cards(
        self, detail_rows: list[tuple[str, str, str, str]]
    ) -> None:
        while self.detail_cards_layout.count():
            child = self.detail_cards_layout.takeAt(0)
            if child.widget() is not None:
                child.widget().deleteLater()
        for fee, amount, invoice, source_row in detail_rows:
            card = QFrame()
            card.setObjectName("rpaFeeCard")
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(10, 8, 10, 8)
            card_layout.setSpacing(5)
            top = QHBoxLayout()
            fee_label = QLabel(fee)
            fee_label.setWordWrap(True)
            fee_label.setStyleSheet("font-weight: 700; color: #15304E;")
            top.addWidget(fee_label, 1)
            amount_label = QLabel(f"{amount} đ" if amount != "—" else "—")
            amount_label.setStyleSheet("font-weight: 700; color: #154999;")
            top.addWidget(amount_label)
            card_layout.addLayout(top)
            invoice_label = QLabel(f"Số HĐ: {invoice}  ·  Dòng BK: {source_row}")
            invoice_label.setWordWrap(True)
            invoice_label.setStyleSheet("color: #5B6F86;")
            card_layout.addWidget(invoice_label)
            self.detail_cards_layout.addWidget(card)
        self.detail_cards_layout.addStretch()

    def resizeEvent(self, event: Any) -> None:
        super().resizeEvent(event)
        self._update_responsive_layout()

    def _toggle_detail(self) -> None:
        if self.width() >= 1180:
            self._detail_user_hidden = not self._detail_user_hidden
            self._update_responsive_layout()
            return
        if self._detail_popup is None:
            popup = QDialog(self)
            popup.setWindowTitle("Chi tiết phí và số HĐ")
            popup.setMinimumSize(470, 390)
            popup_layout = QVBoxLayout(popup)
            self._popup_title = QLabel()
            self._popup_title.setWordWrap(True)
            popup_layout.addWidget(self._popup_title)
            self._popup_table = QTableWidget(0, 4)
            self._popup_table.setHorizontalHeaderLabels(
                ["Loại phí", "Số tiền", "Số HĐ", "Dòng BK"]
            )
            self._popup_table.setEditTriggers(
                QAbstractItemView.EditTrigger.NoEditTriggers
            )
            self._popup_table.setSelectionMode(
                QAbstractItemView.SelectionMode.NoSelection
            )
            self._popup_table.verticalHeader().setVisible(False)
            self._popup_table.horizontalHeader().setSectionResizeMode(
                QHeaderView.ResizeMode.Stretch
            )
            popup_layout.addWidget(self._popup_table, 1)
            close_button = QPushButton("Đóng")
            close_button.clicked.connect(popup.close)
            popup_layout.addWidget(close_button, 0, Qt.AlignmentFlag.AlignRight)
            self._detail_popup = popup
        self._sync_detail_popup()
        self._detail_popup.show()
        self._detail_popup.raise_()
        self._detail_popup.activateWindow()

    def _sync_detail_popup(self) -> None:
        if self._detail_popup is None:
            return
        self._popup_title.setText(self.detail_title.text())
        self._popup_table.setRowCount(self.detail_table.rowCount())
        for row in range(self.detail_table.rowCount()):
            for column in range(self.detail_table.columnCount()):
                original = self.detail_table.item(row, column)
                if original is not None:
                    self._popup_table.setItem(row, column, original.clone())

    def _add_item(self, source: Any) -> None:
        if self._multiple_sheets:
            self._add_compact_item(source)
            return
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
        is_latest_bk = str(source.sqt) in self._latest_bk_sqt
        is_latest = str(source.sqt) in self._latest_pad_sqt
        check.setData(LATEST_BK_ROLE, is_latest_bk)
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
            if is_latest_bk and is_latest
            else "Vừa ghi BK"
            if is_latest_bk
            else "Lượt nhập gần nhất"
            if is_latest
            else ""
        )
        sqt_sort = self._sqt_sort_key(source.sqt)
        group_sort = (
            0
            if is_latest_bk and is_latest
            else 1
            if is_latest_bk
            else 2
            if is_latest
            else 3
        )
        source_rows = tuple(int(value) for value in source.source_rows)
        invoice_numbers = tuple(getattr(source, "invoice_numbers", ()) or ())
        invoice_display = ", ".join(invoice_numbers)
        values = (
            source.sqt,
            invoice_display,
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
            invoice_display.casefold(),
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
        if self._multiple_sheets:
            values += (source.sheet_name,)
            sort_values += (source.sheet_name.casefold(),)
        for column, (value, sort_value) in enumerate(
            zip(values, sort_values), 1
        ):
            cell = _SortableTableWidgetItem(str(value), sort_value)
            cell.setToolTip(str(value))
            if column == 3:
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
            elif column == 4 and group_display:
                font = cell.font()
                font.setBold(True)
                cell.setFont(font)
                cell.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if is_latest_bk and is_latest:
                    cell.setForeground(QColor("#15803D"))
                    cell.setBackground(QColor("#DCFCE7"))
                elif is_latest_bk:
                    cell.setForeground(QColor("#1D4ED8"))
                    cell.setBackground(QColor("#DBEAFE"))
                else:
                    cell.setForeground(QColor("#6D28D9"))
                    cell.setBackground(QColor("#EDE9FE"))
            if not source.can_run:
                cell.setForeground(QColor("#8A3B32"))
            elif str(source.sqt) in self._skipped_sqt and column != 3:
                cell.setBackground(QColor("#FFF4CC"))
            self.table.setItem(row, column, cell)

    @staticmethod
    def _invoice_tokens(source: Any) -> list[str]:
        grouped: OrderedDict[str, list[str]] = OrderedDict()
        for entry in getattr(source, "fee_entries", ()):
            invoice = str(entry.invoice_number or "").strip()
            if not invoice:
                continue
            label = str(entry.fee_label)
            values = grouped.setdefault(label, [])
            if invoice.casefold() not in {value.casefold() for value in values}:
                values.append(invoice)
        if not grouped:
            invoices = tuple(getattr(source, "invoice_numbers", ()) or ())
            if invoices:
                grouped["HĐ chưa gắn phí"] = list(invoices)
        return [f"{label}: {', '.join(values)}" for label, values in grouped.items()]

    @classmethod
    def _invoice_summary(cls, source: Any, search_text: str = "") -> str:
        tokens = cls._invoice_tokens(source)
        if search_text:
            tokens.sort(key=lambda value: search_text not in value.casefold())
        if not tokens:
            return "Chưa có số HĐ trong BK"
        return "\n".join(
            "  ·  ".join(tokens[index:index + 3])
            for index in range(0, len(tokens), 3)
        )

    def _add_compact_item(self, source: Any) -> None:
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setRowHeight(row, 54)
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
        check.setData(SOURCE_ROLE, source)
        check.setData(STATUS_ROLE, source.status)
        check.setData(LATEST_BK_ROLE, str(source.sqt) in self._latest_bk_sqt)
        check.setData(LATEST_PAD_ROLE, str(source.sqt) in self._latest_pad_sqt)
        check.setToolTip(source.validation_message or "Chọn để gửi SQT sang PAD.")
        if str(source.sqt) in self._skipped_sqt:
            check.setBackground(QColor("#FFF4CC"))
        self._checks.append(check)
        self.table.setItem(row, 0, check)

        sqt_sort = self._sqt_sort_key(source.sqt)
        status_sort = (1 if source.status == RPA_STATUS_IMPORTED else 0, sqt_sort)
        values = (
            (str(source.sqt), sqt_sort),
            (source.sheet_name, source.sheet_name.casefold()),
            (self._invoice_summary(source), str(source.sqt)),
            (
                "✓ Đã nhập" if source.status == RPA_STATUS_IMPORTED else "● Chưa nhập",
                status_sort,
            ),
            (f"{_money(source.amounts.total)} đ", source.amounts.total),
        )
        for column, (value, sort_value) in enumerate(values, 1):
            cell = _SortableTableWidgetItem(value, sort_value)
            cell.setToolTip(
                value if column != 3 else self._invoice_summary(source)
            )
            if column == 3:
                cell.setForeground(QColor("#174A75"))
            elif column == 4:
                cell.setForeground(
                    QColor("#15803D") if source.status == RPA_STATUS_IMPORTED
                    else QColor("#1D4ED8")
                )
                cell.setData(STATUS_ROLE, source.status)
            if not source.can_run:
                cell.setForeground(QColor("#8A3B32"))
            self.table.setItem(row, column, cell)

    def _clear_remembered(self) -> None:
        if self._clear_saved_callback is None:
            return
        self._clear_saved_callback()
        self._initial_selected_sqt.clear()
        self._set_all(Qt.CheckState.Unchecked, visible_only=False)
        if hasattr(self, "restore_label"):
            self.restore_label.setText("Đã xóa lựa chọn ghi nhớ của file BK này.")

    def _set_all(
        self, state: Qt.CheckState, *, visible_only: bool = True
    ) -> None:
        self.table.blockSignals(True)
        try:
            for item in self._checks:
                if (
                    item.flags() & Qt.ItemFlag.ItemIsUserCheckable
                    and (not visible_only or not self.table.isRowHidden(item.row()))
                ):
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
        if self._multiple_sheets:
            self._apply_compact_filter()
            return
        search_text = (
            self.sqt_search.text().strip().casefold()
            if hasattr(self, "sqt_search")
            else ""
        )
        for row in range(self.table.rowCount()):
            check = self.table.item(row, 0)
            status = check.data(STATUS_ROLE) if check is not None else None
            latest_bk = bool(check.data(LATEST_BK_ROLE)) if check is not None else False
            latest = bool(check.data(LATEST_PAD_ROLE)) if check is not None else False
            group_matches = (
                bool(search_text)
                or self._active_group_filter == "all"
                or (self._active_group_filter == "bk" and latest_bk)
                or (self._active_group_filter == "latest" and latest)
            )
            status_matches = (
                bool(search_text)
                or self._active_status_filter is None
                or status == self._active_status_filter
            )
            sqt_item = self.table.item(row, 1)
            invoice_item = self.table.item(row, 2)
            sheet_item = (
                self.table.item(row, len(self.COLUMNS))
                if self._multiple_sheets else None
            )
            search_matches = (
                not search_text
                or (
                    sqt_item is not None
                    and search_text in sqt_item.text().strip().casefold()
                )
                or (
                    invoice_item is not None
                    and search_text in invoice_item.text().strip().casefold()
                )
                or (
                    sheet_item is not None
                    and search_text in sheet_item.text().strip().casefold()
                )
            )
            self.table.setRowHidden(
                row,
                not (group_matches and status_matches and search_matches),
            )
        self._update_selection_summary()

    def _apply_compact_filter(self) -> None:
        search_text = self.sqt_search.text().strip().casefold()
        selected_sheet = self.sheet_filter.currentData()
        visible_rows: list[int] = []
        invoice_match_rows: list[int] = []
        for row in range(self.table.rowCount()):
            check = self.table.item(row, 0)
            source = check.data(SOURCE_ROLE) if check is not None else None
            if source is None:
                continue
            latest_bk = bool(check.data(LATEST_BK_ROLE))
            latest_pad = bool(check.data(LATEST_PAD_ROLE))
            group_matches = (
                self._active_group_filter == "all"
                or (self._active_group_filter == "bk" and latest_bk)
                or (self._active_group_filter == "latest" and latest_pad)
            )
            invoice_matches = bool(search_text) and any(
                search_text in str(value).casefold()
                for value in source.invoice_numbers
            )
            search_matches = (
                not search_text
                or search_text in str(source.sqt).casefold()
                or search_text in str(source.sheet_name).casefold()
                or invoice_matches
            )
            visible = (
                group_matches
                and (self._active_status_filter is None
                     or source.status == self._active_status_filter)
                and (selected_sheet is None or source.sheet_name == selected_sheet)
                and search_matches
            )
            self.table.setRowHidden(row, not visible)
            if visible:
                visible_rows.append(row)
                if invoice_matches:
                    invoice_match_rows.append(row)
            invoice_cell = self.table.item(row, 3)
            if invoice_cell is not None:
                invoice_cell.setText(self._invoice_summary(source, search_text))
                invoice_cell.setBackground(
                    QColor("#FFF4CC") if invoice_matches
                    else QColor("#FFFFFF")
                )
        if visible_rows:
            current = self.table.currentRow()
            preferred = (
                invoice_match_rows[0] if invoice_match_rows and search_text
                else current if current in visible_rows else visible_rows[0]
            )
            if current != preferred:
                self.table.setCurrentCell(preferred, 1)
            else:
                self._show_detail()
        else:
            self.table.clearSelection()
            self._show_detail()
        self._update_selection_summary()

    def _show_detail(self, *_args: Any) -> None:
        if not self._multiple_sheets or not hasattr(self, "detail_table"):
            return
        row = self.table.currentRow()
        check = self.table.item(row, 0) if row >= 0 else None
        source = check.data(SOURCE_ROLE) if check is not None else None
        if source is None or self.table.isRowHidden(row):
            self.detail_title.setText("Chọn một dòng SQT để xem phí và HĐ theo dòng BK.")
            self.detail_table.setRowCount(0)
            self._refresh_detail_cards([])
            if hasattr(self, "_detail_popup"):
                self._sync_detail_popup()
            return
        self.detail_title.setText(
            f"SQT {source.sqt} · {source.sheet_name} · "
            f"Tổng khoản chi {_money(source.amounts.total)} đ"
            + (f" · {source.validation_message}" if source.errors else "")
        )
        category_labels = {
            key: label for key, label, _fees in RPA_FEE_GROUPS
        }
        fee_order = {
            fee_key: (group_index, fee_index)
            for group_index, (_key, _label, fees) in enumerate(RPA_FEE_GROUPS)
            for fee_index, (fee_key, _fee_label) in enumerate(fees)
        }
        entries = sorted(
            getattr(source, "fee_entries", ()),
            key=lambda entry: (
                *fee_order.get(entry.fee_key, (99, 99)), entry.source_row
            ),
        )
        detail_rows: list[tuple[str, str, str, str]] = []
        for entry in entries:
            category = category_labels.get(entry.category_key, "")
            label = (
                entry.fee_label if not category or category == entry.fee_label
                else f"{category} / {entry.fee_label}"
            )
            invoice = entry.invoice_number or (
                "Chưa có HĐ" if entry.has_invoice_column else "Không có cột HĐ"
            )
            detail_rows.append((
                label,
                _money(entry.amount) if entry.amount is not None else "—",
                invoice,
                str(entry.source_row),
            ))
        covered_categories = {entry.category_key for entry in entries}
        for key, label, _fees in RPA_FEE_GROUPS:
            amount = source.amounts.to_dict().get(key, 0)
            if amount and key not in covered_categories:
                detail_rows.append((
                    label, _money(amount), "HĐ chưa xác định",
                    ", ".join(str(value) for value in source.source_rows),
                ))
        if not detail_rows:
            detail_rows.append(("Chưa có phí", "0", "Chưa có HĐ", "—"))
        self._refresh_detail_cards(detail_rows)
        self.detail_table.setRowCount(len(detail_rows))
        search_text = self.sqt_search.text().strip().casefold()
        first_match: QTableWidgetItem | None = None
        for detail_row, values in enumerate(detail_rows):
            for column, value in enumerate(values):
                cell = QTableWidgetItem(value)
                cell.setToolTip(value)
                if column == 2 and search_text and search_text in value.casefold():
                    cell.setBackground(QColor("#FFF4CC"))
                    first_match = first_match or cell
                self.detail_table.setItem(detail_row, column, cell)
        if first_match is not None:
            self.detail_table.scrollToItem(first_match)
        if hasattr(self, "_detail_popup"):
            self._sync_detail_popup()

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
            if self._multiple_sheets:
                self.selection_summary.setText(
                    f"Đã chọn {len(selected)} SQT · "
                    f"đang hiển thị {visible_selected} · "
                    f"{not_imported} chưa nhập, {imported} nhập lại"
                )
                self.footer_summary.setText(
                    f"{len(selected)} SQT đã chọn · {_money(total_amount)} đ"
                )
            else:
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
            if self._multiple_sheets:
                self.footer_summary.setText("Chưa chọn SQT nào")
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
        items = self.payload.get("items")
        self._sheet_names = {
            str(item.get("sheet_name") or self.payload.get("sheet_name") or "")
            for item in items if isinstance(item, Mapping)
        } if isinstance(items, list) else set()
        self._multiple_sheets = len(self._sheet_names) > 1
        self.setObjectName("rpaLatestDataDialog")
        self.setWindowTitle("Dữ liệu của lần nhập gần nhất")
        self._build_ui()
        fit_window_to_screen(self, 1200, 650)

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        sheet_name = (
            "nhiều sheet BK"
            if self._multiple_sheets
            else str(
                self.payload.get("sheet_name")
                or next(iter(self._sheet_names), "")
                or "—"
            )
        )
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
        columns = self.COLUMNS + (("Sheet BK",) if self._multiple_sheets else ())
        self.table = QTableWidget(0, len(columns))
        self.table.setObjectName("rpaLatestDataTable")
        self.table.setHorizontalHeaderLabels(list(columns))
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
        if self._multiple_sheets:
            cells += (source.get("sheet_name") or "—",)
        row = self.table.rowCount()
        self.table.insertRow(row)
        for column, value in enumerate(cells):
            item = QTableWidgetItem(str(value if value is not None else "—"))
            item.setToolTip(item.text())
            self.table.setItem(row, column, item)


__all__ = ["RpaLatestDataDialog", "RpaSqtSelectionDialog"]
