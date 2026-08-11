"""Dialog chọn sheet tháng và xử lý tập trung xung đột Excel."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from datetime import datetime
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


VALID_FEE_CODES = (
    "CB",
    "CBDH",
    "VTN",
    "NV",
    "HH",
    "NH",
    "HV",
    "VSDL",
    "LC",
    "QT",
    "LL",
    "SC",
)


def _value(source: Any, *names: str, default: Any = None) -> Any:
    if source is None:
        return default
    if isinstance(source, Mapping):
        for name in names:
            if name in source:
                return source[name]
        return default
    for name in names:
        if hasattr(source, name):
            value = getattr(source, name)
            return getattr(value, "value", value)
    return default


def _sequence(value: Any) -> tuple[Any, ...]:
    if value is None:
        return ()
    if isinstance(value, Mapping):
        return tuple(value.values())
    if isinstance(value, Sequence) and not isinstance(
        value, (str, bytes, bytearray)
    ):
        return tuple(value)
    try:
        return tuple(value)
    except TypeError:
        return (value,)


def _code(value: Any) -> str:
    raw = _value(value, "value", "code", "action", "id", default=value)
    return str(getattr(raw, "value", raw) or "").split(".")[-1].upper()


def _display(value: Any) -> str:
    if value in (None, ""):
        return "—"
    raw = getattr(value, "value", value)
    return str(raw)


def _format_amount(value: Any) -> str:
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return f"{value:,}".replace(",", ".")
    if isinstance(value, float) and value.is_integer():
        return f"{int(value):,}".replace(",", ".")
    return _display(value)


def _format_datetime(value: Any) -> str:
    if value in (None, ""):
        return "—"
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y %H:%M:%S")
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00")).strftime(
            "%d/%m/%Y %H:%M:%S"
        )
    except (TypeError, ValueError):
        return str(value)


def _sheet_name(candidate: Any) -> str:
    return str(
        _value(
            candidate,
            "sheet_name",
            "name",
            "target_sheet",
            "sheet",
            default="",
        )
        or ""
    )


class MonthSelectionDialog(QDialog):
    """Chọn đúng một sheet khi analyze tìm thấy nhiều tháng phù hợp."""

    def __init__(
        self,
        candidates_or_plan: Any,
        parent: QWidget | None = None,
        *,
        title: str = "Chọn tháng xử lý",
        preselect_first: bool = True,
        show_recommendations: bool = True,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("monthSelectionDialog")
        self.setWindowTitle(title)
        self.resize(650, 390)
        candidates = _value(
            candidates_or_plan,
            "month_candidates",
            "sheet_candidates",
            "target_sheet_candidates",
            default=candidates_or_plan,
        )
        self.candidates = list(_sequence(candidates))
        self.preselect_first = preselect_first
        self.show_recommendations = show_recommendations
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        note = QLabel(
            "Hãy chọn một sheet để tiếp tục; "
            "chỉ sheet này được phép thay đổi."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        self.table = QTableWidget(0, 5)
        self.table.setObjectName("monthCandidateTable")
        self.table.setHorizontalHeaderLabels(
            ["Sheet", "Tháng", "Năm", "Cập nhật / mới", "Gần nhất"]
        )
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        if not self.show_recommendations:
            self.table.setColumnHidden(3, True)
            self.table.setColumnHidden(4, True)
        layout.addWidget(self.table, 1)

        for candidate in self.candidates:
            row = self.table.rowCount()
            self.table.insertRow(row)
            sheet_item = QTableWidgetItem(_sheet_name(candidate))
            sheet_item.setData(Qt.ItemDataRole.UserRole, candidate)
            self.table.setItem(row, 0, sheet_item)
            self.table.setItem(
                row,
                1,
                QTableWidgetItem(
                    _display(_value(candidate, "month", "month_number"))
                ),
            )
            self.table.setItem(
                row,
                2,
                QTableWidgetItem(_display(_value(candidate, "year"))),
            )
            update_count = _value(candidate, "update_count", default=None)
            if update_count is not None and _value(
                candidate, "source_sheet", default=""
            ):
                count = (
                    f"{int(update_count or 0)} / "
                    f"{int(_value(candidate, 'new_row_count', default=0) or 0)}"
                )
            else:
                count = _value(
                    candidate,
                    "new_row_count",
                    "match_count",
                    "row_count",
                    "count",
                )
            self.table.setItem(row, 3, QTableWidgetItem(_display(count)))
            recent = bool(
                _value(
                    candidate,
                    "is_recent",
                    "recent",
                    "recently_synced",
                    default=False,
                )
            )
            self.table.setItem(
                row,
                4,
                QTableWidgetItem("Vừa đồng bộ" if recent else ""),
            )

        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Tiếp tục")
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Hủy")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self.table.itemSelectionChanged.connect(self._update_action)
        self.table.itemDoubleClicked.connect(lambda _item: self.accept())
        if self.candidates and self.preselect_first:
            self.table.selectRow(0)
        self._update_action()

    def _update_action(self) -> None:
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(
            self.table.currentRow() >= 0
        )

    @property
    def selected_candidate(self) -> Any | None:
        row = self.table.currentRow()
        item = self.table.item(row, 0) if row >= 0 else None
        return item.data(Qt.ItemDataRole.UserRole) if item is not None else None

    @property
    def selected_sheet_name(self) -> str | None:
        candidate = self.selected_candidate
        return _sheet_name(candidate) or None

    def selection(self) -> dict[str, Any]:
        candidate = self.selected_candidate
        sheet_name = self.selected_sheet_name
        month = _value(candidate, "month", "month_number")
        return {
            "sheet_name": sheet_name,
            "selected_sheet_name": sheet_name,
            "selected_sheet": sheet_name,
            "month": month,
            "selected_month": month,
            "candidate": candidate,
        }

    resolution = selection


class ManualRowPickerDialog(QDialog):
    """Chọn một dòng BK trong cửa sổ ba tháng, không cho phép chọn cột."""

    def __init__(
        self,
        candidates: Any,
        parent: QWidget | None = None,
        *,
        sheet_name: str = "",
    ) -> None:
        super().__init__(parent)
        self.candidates = list(_sequence(candidates))
        self.setObjectName("manualRowPickerDialog")
        self.setWindowTitle(
            f"Chọn dòng trong {sheet_name}" if sheet_name else "Chọn dòng BK"
        )
        self.resize(850, 410)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        note = QLabel(
            "Chọn đúng sheet nguồn và dòng kế hoạch. "
            "Cột khoản chi vẫn do tiêu đề workbook quyết định."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        self.table = QTableWidget(0, 7)
        self.table.setObjectName("manualRowCandidateTable")
        self.table.setHorizontalHeaderLabels(
            [
                "Sheet nguồn",
                "SQT",
                "Container",
                "Loại hàng",
                "Ngày đóng",
                "Tàu",
                "Người nhận",
            ]
        )
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(True)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table, 1)

        for candidate in self.candidates:
            row = self.table.rowCount()
            self.table.insertRow(row)
            values = (
                _value(candidate, "source_sheet", "sheet_name", "sheet"),
                _value(candidate, "sqt", "sequence_number"),
                _value(candidate, "container", "container_number"),
                _value(candidate, "cargo_type", "goods_type", "loai_hang"),
                _value(candidate, "closing_date", "ngay_dong"),
                _value(candidate, "vessel", "ship", "ten_tau"),
                _value(candidate, "recipient", "nguoi_nhan"),
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(_display(value))
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, candidate)
                self.table.setItem(row, column, item)
            workbook_row = _value(
                candidate,
                "row",
                "row_number",
                "target_row",
                "excel_row",
                default=(candidate if isinstance(candidate, int) else row + 1),
            )
            self.table.setVerticalHeaderItem(
                row, QTableWidgetItem(str(workbook_row))
            )

        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Chọn dòng")
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Hủy")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self.table.itemSelectionChanged.connect(self._update_action)
        self.table.itemDoubleClicked.connect(lambda _item: self.accept())
        if self.candidates:
            self.table.selectRow(0)
        self._update_action()

    def _update_action(self) -> None:
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(
            self.table.currentRow() >= 0
        )

    @property
    def selected_candidate(self) -> Any | None:
        row = self.table.currentRow()
        item = self.table.item(row, 0) if row >= 0 else None
        return item.data(Qt.ItemDataRole.UserRole) if item is not None else None

    @property
    def selected_row(self) -> Any | None:
        candidate = self.selected_candidate
        if candidate is None:
            return None
        return _value(
            candidate,
            "row",
            "row_number",
            "target_row",
            "excel_row",
            default=(candidate if isinstance(candidate, int) else None),
        )

    @property
    def selected_source_sheet(self) -> str | None:
        candidate = self.selected_candidate
        value = _value(candidate, "source_sheet", "sheet_name", "sheet")
        return str(value) if value not in (None, "") else None


DEFAULT_ACTIONS: dict[str, tuple[str, ...]] = {
    "INVALID_SQT": ("SKIP_INVALID", "CANCEL_ALL"),
    "DUPLICATE": ("KEEP_ONE", "KEEP_ALL", "CANCEL_ALL"),
    "DUPLICATE_SYNC_ROW": ("KEEP_ONE", "KEEP_ALL", "CANCEL_ALL"),
    "DUPLICATE_SOURCE_ROW": ("KEEP_ONE", "KEEP_ALL", "CANCEL_ALL"),
    "SYNC_GROUP_COUNT_MISMATCH": ("CANCEL_ALL",),
    "TARGET_MONTH_AMBIGUOUS": ("SELECT_MONTH", "CANCEL"),
    "TARGET_SHEET_AMBIGUOUS": ("SELECT_SHEET", "CANCEL"),
    "CONTAINER_NOT_FOUND": ("SKIP", "SELECT_ROW"),
    "MULTIPLE_CONTAINER_MATCH": ("SKIP", "SELECT_ROW"),
    "REPEATED_SOURCE_CONTAINER": ("SKIP", "SELECT_ROW"),
    "TARGET_CELL_OCCUPIED": ("KEEP_EXISTING", "OVERWRITE", "SKIP"),
    "TARGET_CELL_FORMULA": ("KEEP_FORMULA", "OVERWRITE", "SKIP"),
    "TARGET_CELL_TEXT": ("KEEP_EXISTING", "OVERWRITE", "SKIP"),
    "UNKNOWN_FEE_CODE": ("SELECT_FEE", "SKIP"),
    "FEE_COLUMN_MISSING": ("SKIP", "CANCEL_ALL"),
    "BL_ONLY_NO_CONTAINER": ("SKIP", "SELECT_ROW"),
    "PARTIAL_KEY_MATCH": ("SKIP", "SELECT_ROW", "CANCEL_ALL"),
    "PAYMENT_SOURCE_INVALID": ("SKIP", "CANCEL_ALL"),
    "PAYMENT_CLEAR_VALUE": (
        "KEEP_EXISTING",
        "OVERWRITE",
        "SKIP",
        "CANCEL_ALL",
    ),
    "MULTIPLE_SOURCE_INVOICES": ("SELECT_INVOICE",),
    "INVOICE_VALUE_CONFLICT": ("KEEP_EXISTING", "OVERWRITE"),
    "INVOICE_COLUMN_MISSING": ("SKIP_INVOICE", "CANCEL_ALL"),
    "MULTIPLE_EXPENSE_SAME_CELL": ("SELECT_SOURCE_ITEM",),
    "CARRY_FORWARD_MAPPING_INVALID": ("SKIP", "CANCEL_ALL"),
    "MULTIPLE_SOURCE_CARRIERS": ("SELECT_CARRIER",),
    "CARRIER_VALUE_CONFLICT": (
        "KEEP_EXISTING",
        "APPEND_CARRIER",
        "OVERWRITE",
    ),
    "CARRIER_COLUMN_MISSING": ("KEEP_EXISTING", "CANCEL_ALL"),
    "BATCH_ALREADY_POSTED": ("POST_UNPOSTED_ONLY", "CANCEL"),
    "FILE_CHANGED": ("REANALYZE", "CANCEL"),
    "FILE_LOCKED": ("RETRY", "CANCEL"),
}

DEFAULT_RESOLUTION: dict[str, str] = {
    "INVALID_SQT": "SKIP_INVALID",
    "DUPLICATE": "KEEP_ONE",
    "DUPLICATE_SYNC_ROW": "KEEP_ONE",
    "DUPLICATE_SOURCE_ROW": "KEEP_ONE",
    "SYNC_GROUP_COUNT_MISMATCH": "CANCEL_ALL",
    "TARGET_MONTH_AMBIGUOUS": "SELECT_MONTH",
    "TARGET_SHEET_AMBIGUOUS": "SELECT_SHEET",
    "CONTAINER_NOT_FOUND": "SKIP",
    "MULTIPLE_CONTAINER_MATCH": "SKIP",
    "REPEATED_SOURCE_CONTAINER": "SKIP",
    "TARGET_CELL_OCCUPIED": "KEEP_EXISTING",
    "TARGET_CELL_FORMULA": "KEEP_FORMULA",
    "TARGET_CELL_TEXT": "KEEP_EXISTING",
    "UNKNOWN_FEE_CODE": "SKIP",
    "FEE_COLUMN_MISSING": "SKIP",
    "BL_ONLY_NO_CONTAINER": "SKIP",
    "PARTIAL_KEY_MATCH": "SKIP",
    "PAYMENT_SOURCE_INVALID": "SKIP",
    "PAYMENT_CLEAR_VALUE": "KEEP_EXISTING",
    "MULTIPLE_SOURCE_INVOICES": "SELECT_INVOICE",
    "INVOICE_VALUE_CONFLICT": "KEEP_EXISTING",
    "INVOICE_COLUMN_MISSING": "SKIP_INVOICE",
    "MULTIPLE_EXPENSE_SAME_CELL": "SELECT_SOURCE_ITEM",
    "CARRY_FORWARD_MAPPING_INVALID": "SKIP",
    "MULTIPLE_SOURCE_CARRIERS": "SELECT_CARRIER",
    "CARRIER_VALUE_CONFLICT": "KEEP_EXISTING",
    "CARRIER_COLUMN_MISSING": "KEEP_EXISTING",
    "BATCH_ALREADY_POSTED": "POST_UNPOSTED_ONLY",
    "FILE_CHANGED": "REANALYZE",
    "FILE_LOCKED": "RETRY",
}

ACTION_LABELS = {
    "SKIP": "Bỏ qua",
    "SKIP_INVOICE": "Bỏ ghi HĐ",
    "SKIP_INVALID": "Bỏ qua dòng lỗi",
    "CANCEL": "Hủy",
    "CANCEL_ALL": "Hủy toàn bộ",
    "KEEP_ONE": "Chỉ giữ một dòng",
    "KEEP_ALL": "Giữ tất cả",
    "SELECT_SHEET": "Chọn sheet",
    "SELECT_MONTH": "Chọn tháng",
    "SELECT_ROW": "Chọn dòng",
    "SELECT_FEE": "Chọn mã phí",
    "SELECT_INVOICE": "Chọn một Số HĐ",
    "SELECT_SOURCE_ITEM": "Chọn một dòng JSON",
    "SELECT_CARRIER": "Chọn bên vận tải",
    "KEEP_EXISTING": "Giữ nguyên",
    "KEEP_FORMULA": "Giữ công thức",
    "OVERWRITE": "Ghi đè",
    "APPEND_CARRIER": "Ghi thêm",
    "POST_UNPOSTED_ONLY": "Chỉ nhập khoản chưa xử lý",
    "REANALYZE": "Đọc lại dữ liệu",
    "RETRY": "Thử lại",
}


class RepostSelectionDialog(QDialog):
    """Chọn các khoản đã nhập trước đây cần được nhập lại."""

    COLUMNS = (
        "Chọn",
        "Dòng JSON",
        "Container",
        "Phí",
        "Số tiền",
        "Sheet",
        "Dòng / ô",
        "Đã nhập lúc",
    )

    def __init__(
        self,
        items: Sequence[Any],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.items = list(items)
        self.setObjectName("repostSelectionDialog")
        self.setWindowTitle("Chọn khoản nhập lại")
        self.resize(920, 520)
        self._checkbox_items: list[QTableWidgetItem] = []
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        note = QLabel(
            f"Có {len(self.items)} khoản đã được nhập trước đây. "
            "Mặc định phần mềm chỉ nhập các khoản chưa nhập."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        self.unposted_only = QRadioButton("Chỉ nhập khoản chưa nhập")
        self.choose_reposts = QRadioButton("Chọn khoản nhập lại")
        self.select_all = QCheckBox("Chọn tất cả")
        self.unposted_only.setChecked(True)
        layout.addWidget(self.unposted_only)
        layout.addWidget(self.choose_reposts)
        layout.addWidget(self.select_all)

        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setObjectName("repostItemTable")
        self.table.setHorizontalHeaderLabels(list(self.COLUMNS))
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table, 1)

        for source in self.items:
            row = self.table.rowCount()
            self.table.insertRow(row)
            checkbox = QTableWidgetItem()
            checkbox.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsSelectable
            )
            checkbox.setCheckState(Qt.CheckState.Unchecked)
            checkbox.setData(
                Qt.ItemDataRole.UserRole,
                _value(source, "source_item_index"),
            )
            self._checkbox_items.append(checkbox)
            self.table.setItem(row, 0, checkbox)
            row_cell = " / ".join(
                part
                for part in (
                    _display(_value(source, "target_row")),
                    _display(_value(source, "target_cell")),
                )
                if part != "—"
            )
            values = (
                (
                    int(_value(source, "source_item_index", default=0)) + 1
                ),
                _value(source, "container"),
                _value(source, "fee", "fee_selected"),
                _format_amount(_value(source, "amount")),
                _value(source, "sheet_name"),
                row_cell or "—",
                _format_datetime(_value(source, "created_at")),
            )
            for column, value in enumerate(values, 1):
                self.table.setItem(row, column, QTableWidgetItem(_display(value)))

        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Tiếp tục")
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Hủy")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self.choose_reposts.toggled.connect(self._update_enabled)
        self.select_all.stateChanged.connect(self._set_all_checked)
        self.table.itemChanged.connect(self._sync_select_all_state)
        self._update_enabled(False)

    def _update_enabled(self, enabled: bool) -> None:
        self.table.setEnabled(enabled)
        self.select_all.setEnabled(enabled)

    def _set_all_checked(self, state: int) -> None:
        check_state = Qt.CheckState(state)
        if check_state is Qt.CheckState.PartiallyChecked:
            return
        target_state = (
            Qt.CheckState.Checked
            if check_state is Qt.CheckState.Checked
            else Qt.CheckState.Unchecked
        )
        was_blocked = self.table.blockSignals(True)
        try:
            for item in self._checkbox_items:
                item.setCheckState(target_state)
        finally:
            self.table.blockSignals(was_blocked)

    def _sync_select_all_state(self, changed_item: QTableWidgetItem) -> None:
        if changed_item.column() != 0:
            return
        checked_count = sum(
            item.checkState() is Qt.CheckState.Checked
            for item in self._checkbox_items
        )
        if checked_count == len(self._checkbox_items) and self._checkbox_items:
            state = Qt.CheckState.Checked
        elif checked_count:
            state = Qt.CheckState.PartiallyChecked
        else:
            state = Qt.CheckState.Unchecked
        was_blocked = self.select_all.blockSignals(True)
        try:
            self.select_all.setCheckState(state)
        finally:
            self.select_all.blockSignals(was_blocked)

    @property
    def selected_source_indices(self) -> list[int]:
        if not self.choose_reposts.isChecked():
            return []
        return [
            int(item.data(Qt.ItemDataRole.UserRole))
            for item in self._checkbox_items
            if item.checkState() == Qt.CheckState.Checked
        ]


class ConflictResolutionDialog(QDialog):
    """Một bảng duy nhất để giải quyết toàn bộ xung đột của một plan."""

    COLUMNS = (
        "Container",
        "B/L",
        "SQT",
        "Phí",
        "Số tiền",
        "Sheet",
        "Dòng",
        "Cột / ô",
        "Số HĐ từ JSON",
        "Giá trị hiện tại",
        "Vấn đề",
        "Cách xử lý",
        "Dòng / phí / sheet chọn",
    )

    def __init__(
        self,
        plan_or_conflicts: Any,
        parent: QWidget | None = None,
        *,
        valid_fee_codes: Sequence[str] = VALID_FEE_CODES,
    ) -> None:
        super().__init__(parent)
        self.plan = plan_or_conflicts
        raw_conflicts = _value(
            plan_or_conflicts,
            "conflicts",
            "unresolved_conflicts",
            default=plan_or_conflicts,
        )
        self.conflicts = list(_sequence(raw_conflicts))
        self.valid_fee_codes = tuple(valid_fee_codes)
        self._action_combos: dict[str, QComboBox] = {}
        self._selected_rows: dict[str, Any] = {}
        self._selected_source_sheets: dict[str, str] = {}
        self._selected_source_items: dict[str, QComboBox] = {}
        self._selected_fees: dict[str, QComboBox] = {}
        self._selected_sheets: dict[str, QComboBox] = {}
        self._selected_months: dict[str, QComboBox] = {}
        self._selected_invoices: dict[str, QComboBox] = {}
        self._selected_carriers: dict[str, QComboBox] = {}
        self._selector_buttons: dict[str, QPushButton] = {}
        self.setObjectName("excelConflictResolutionDialog")
        self.setWindowTitle("Xử lý xung đột Excel")
        self.resize(1320, 660)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        title = QLabel(
            f"Có {len(self.conflicts)} mục cần xử lý trước khi ghi workbook."
        )
        title.setStyleSheet("font-size: 12pt; font-weight: 700;")
        layout.addWidget(title)
        note = QLabel(
            "Kiểm tra giá trị hiện tại và chọn cách xử lý cho từng mục. "
            "Workbook chỉ được ghi sau khi toàn bộ lựa chọn hợp lệ."
        )
        note.setWordWrap(True)
        note.setProperty("muted", True)
        layout.addWidget(note)

        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setObjectName("excelConflictTable")
        self.table.setHorizontalHeaderLabels(list(self.COLUMNS))
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.table.horizontalHeader().setSectionResizeMode(
            10, QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self.table, 1)

        for index, conflict in enumerate(self.conflicts):
            self._add_conflict(index, conflict)

        self.validation_label = QLabel()
        self.validation_label.setObjectName("conflictValidationLabel")
        self.validation_label.setProperty("status", "error")
        self.validation_label.setWordWrap(True)
        layout.addWidget(self.validation_label)

        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText(
            "Áp dụng lựa chọn"
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Hủy")
        self.buttons.accepted.connect(self._validate_and_accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

    def _add_conflict(self, row: int, conflict: Any) -> None:
        self.table.insertRow(row)
        conflict_id = self._conflict_id(conflict, row)
        target_cell = _value(conflict, "target_cell", "cell", "cell_address")
        target_column = _value(
            conflict, "target_column", "column", "column_name"
        )
        column_cell = " / ".join(
            part
            for part in (_display(target_column), _display(target_cell))
            if part != "—"
        )
        conflict_type = _code(
            _value(conflict, "conflict_type", "type", "kind", default="CONFLICT")
        )
        details = _value(conflict, "details", default={})
        invoice_candidates = _sequence(
            _value(details, "invoice_candidates", default=())
        )
        invoice_display = ", ".join(str(value) for value in invoice_candidates)
        carrier_candidates = _sequence(
            _value(details, "carrier_candidates", default=())
        )
        if carrier_candidates:
            carrier_display = ", ".join(str(value) for value in carrier_candidates)
            invoice_display = (
                f"{invoice_display} | VT: {carrier_display}"
                if invoice_display
                else f"VT: {carrier_display}"
            )
        values = (
            _value(conflict, "container", "container_number"),
            _value(conflict, "bl", "bill_of_lading"),
            _value(conflict, "sqt", "sequence_number"),
            _value(conflict, "selected_fee", "fee", "fee_code"),
            _format_amount(_value(conflict, "amount", "incoming_amount")),
            _value(conflict, "sheet_name", "sheet", "target_sheet"),
            _value(conflict, "target_row", "row", "row_number"),
            column_cell,
            invoice_display or None,
            _value(
                conflict,
                "current_value",
                "existing_value",
                "value_before",
            ),
            _value(
                conflict,
                "message",
                "reason",
                "problem",
                default=conflict_type,
            ),
        )
        for column, value in enumerate(values):
            item = QTableWidgetItem(_display(value))
            item.setToolTip(_display(value))
            if column == 0:
                item.setData(Qt.ItemDataRole.UserRole, conflict)
                item.setData(Qt.ItemDataRole.UserRole + 1, conflict_id)
            self.table.setItem(row, column, item)

        action_combo = QComboBox()
        action_combo.setObjectName(f"conflictAction_{row}")
        options = self._actions(conflict, conflict_type)
        default_action = _code(
            _value(conflict, "default_action", "default_resolution")
        ) or DEFAULT_RESOLUTION.get(conflict_type, _code(options[0]))
        for option in options:
            code = _code(option)
            label = ACTION_LABELS.get(
                code,
                str(_value(option, "label", default=code)),
            )
            if conflict_type == "PAYMENT_CLEAR_VALUE":
                label = {
                    "KEEP_EXISTING": "Giữ giá trị Thanh toán",
                    "OVERWRITE": "Xóa theo BK",
                    "SKIP": "Bỏ qua bản ghi",
                    "CANCEL_ALL": "Hủy toàn bộ",
                }.get(code, label)
            elif conflict_type == "INVOICE_VALUE_CONFLICT":
                label = {
                    "KEEP_EXISTING": "Giữ HĐ hiện tại",
                    "OVERWRITE": "Ghi đè bằng HĐ mới",
                }.get(code, label)
            action_combo.addItem(label, code)
        default_index = action_combo.findData(default_action)
        action_combo.setCurrentIndex(max(0, default_index))
        self._action_combos[conflict_id] = action_combo
        self.table.setCellWidget(row, 11, action_combo)

        selector = self._selector_for(row, conflict, conflict_id, conflict_type)
        if selector is not None:
            self.table.setCellWidget(row, 12, selector)
        else:
            self.table.setItem(row, 12, QTableWidgetItem("—"))

    def _selector_for(
        self,
        row: int,
        conflict: Any,
        conflict_id: str,
        conflict_type: str,
    ) -> QWidget | None:
        row_candidates = _sequence(
            _value(
                conflict,
                "row_candidates",
                "candidate_rows",
                "candidates",
                default=(),
            )
        )
        if row_candidates and conflict_type in {
            "CONTAINER_NOT_FOUND",
            "MULTIPLE_CONTAINER_MATCH",
            "REPEATED_SOURCE_CONTAINER",
            "BL_ONLY_NO_CONTAINER",
            "PARTIAL_KEY_MATCH",
        }:
            button = QPushButton("Chọn dòng…")
            button.setObjectName(f"selectConflictRow_{row}")
            button.clicked.connect(
                lambda _checked=False, c=conflict, key=conflict_id, b=button: (
                    self._pick_row(c, key, b)
                )
            )
            self._selector_buttons[conflict_id] = button
            return button

        if conflict_type == "UNKNOWN_FEE_CODE":
            combo = QComboBox()
            combo.setObjectName(f"selectConflictFee_{row}")
            fees = _sequence(
                _value(
                    conflict,
                    "valid_fee_codes",
                    "fee_candidates",
                    default=self.valid_fee_codes,
                )
            )
            for fee in fees:
                combo.addItem(_code(fee), _code(fee))
            self._selected_fees[conflict_id] = combo
            return combo

        details = _value(conflict, "details", default={})
        if conflict_type == "MULTIPLE_EXPENSE_SAME_CELL":
            combo = QComboBox()
            combo.setObjectName(f"selectConflictSourceItem_{row}")
            combo.addItem("— Chọn dòng JSON —", None)
            for option in _sequence(
                _value(details, "source_item_options", default=())
            ):
                source_index = _value(option, "source_item_index")
                amount = _format_amount(_value(option, "amount"))
                invoice = _value(option, "invoice_no")
                label = f"Dòng {int(source_index) + 1}: {amount}"
                if invoice not in (None, ""):
                    label += f" – HĐ {invoice}"
                combo.addItem(label, source_index)
            self._selected_source_items[conflict_id] = combo
            return combo
        if conflict_type == "MULTIPLE_SOURCE_INVOICES":
            combo = QComboBox()
            combo.setObjectName(f"selectConflictInvoice_{row}")
            combo.addItem("— Chọn Số HĐ —", None)
            for invoice in _sequence(
                _value(details, "invoice_candidates", default=())
            ):
                combo.addItem(str(invoice), str(invoice))
            self._selected_invoices[conflict_id] = combo
            return combo
        if conflict_type == "MULTIPLE_SOURCE_CARRIERS" or (
            conflict_type == "CARRIER_VALUE_CONFLICT"
            and bool(
                _value(
                    details,
                    "requires_existing_carrier_selection",
                    default=False,
                )
            )
        ):
            combo = QComboBox()
            combo.setObjectName(f"selectConflictCarrier_{row}")
            combo.addItem("— Chọn bên vận tải —", None)
            candidate_field = (
                "existing_carrier_candidates"
                if conflict_type == "CARRIER_VALUE_CONFLICT"
                else "carrier_candidates"
            )
            for carrier in _sequence(_value(details, candidate_field, default=())):
                combo.addItem(str(carrier), str(carrier))
            self._selected_carriers[conflict_id] = combo
            return combo
        sheet_candidates = _sequence(
            _value(
                conflict,
                "sheet_candidates",
                "target_sheet_candidates",
                "month_candidates",
                default=_value(
                    details,
                    "sheet_candidates",
                    "target_sheet_candidates",
                    "month_candidates",
                    default=(),
                ),
            )
        )
        if conflict_type == "TARGET_MONTH_AMBIGUOUS":
            combo = QComboBox()
            combo.setObjectName(f"selectConflictMonth_{row}")
            for candidate in sheet_candidates:
                month = _value(candidate, "month", "month_number")
                name = _sheet_name(candidate)
                label = name or (f"Tháng {month}" if month is not None else "")
                if label:
                    combo.addItem(label, month)
            self._selected_months[conflict_id] = combo
            return combo
        if sheet_candidates or conflict_type == "TARGET_SHEET_AMBIGUOUS":
            combo = QComboBox()
            combo.setObjectName(f"selectConflictSheet_{row}")
            for candidate in sheet_candidates:
                name = _sheet_name(candidate)
                if name:
                    combo.addItem(name, name)
            self._selected_sheets[conflict_id] = combo
            return combo
        return None

    def _pick_row(
        self,
        conflict: Any,
        conflict_id: str,
        button: QPushButton,
    ) -> None:
        candidates = _value(
            conflict,
            "row_candidates",
            "candidate_rows",
            "candidates",
            default=(),
        )
        dialog = ManualRowPickerDialog(
            candidates,
            self,
            sheet_name=str(
                _value(conflict, "sheet_name", "sheet", "target_sheet", default="")
                or ""
            ),
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        selected_row = dialog.selected_row
        self._selected_rows[conflict_id] = selected_row
        selected_source_sheet = dialog.selected_source_sheet
        if selected_source_sheet:
            self._selected_source_sheets[conflict_id] = selected_source_sheet
        button.setText(
            f"{selected_source_sheet} – dòng {selected_row}"
            if selected_source_sheet
            else f"Dòng {selected_row}"
        )

    @staticmethod
    def _actions(conflict: Any, conflict_type: str) -> tuple[Any, ...]:
        explicit = _sequence(
            _value(
                conflict,
                "allowed_actions",
                "allowed_resolutions",
                "actions",
                "options",
                default=(),
            )
        )
        actions = explicit or DEFAULT_ACTIONS.get(
            conflict_type, ("SKIP", "CANCEL_ALL")
        )
        if conflict_type == "TARGET_CELL_OCCUPIED":
            actions = tuple(action for action in actions if _code(action) != "ADD")
        return actions

    @staticmethod
    def _conflict_id(conflict: Any, row: int) -> str:
        value = _value(conflict, "conflict_id", "id", "key")
        return str(value) if value not in (None, "") else f"conflict-{row}"

    def resolution_map(self) -> dict[str, dict[str, Any]]:
        resolutions: dict[str, dict[str, Any]] = {}
        for row, conflict in enumerate(self.conflicts):
            conflict_id = self._conflict_id(conflict, row)
            action = str(self._action_combos[conflict_id].currentData() or "")
            payload: dict[str, Any] = {
                "conflict_id": conflict_id,
                "action": action,
            }
            if conflict_id in self._selected_rows:
                payload["selected_row"] = self._selected_rows[conflict_id]
                payload["row"] = self._selected_rows[conflict_id]
                payload["selected_source_sheet"] = self._selected_source_sheets.get(
                    conflict_id
                )
            if conflict_id in self._selected_fees:
                fee = self._selected_fees[conflict_id].currentData()
                payload["selected_fee"] = fee
                payload["fee"] = fee
            if conflict_id in self._selected_sheets:
                sheet = self._selected_sheets[conflict_id].currentData()
                payload["selected_sheet_name"] = sheet
                payload["selected_sheet"] = sheet
                payload["sheet_name"] = sheet
            if conflict_id in self._selected_months:
                month = self._selected_months[conflict_id].currentData()
                payload["selected_month"] = month
                payload["month"] = month
            if conflict_id in self._selected_invoices:
                invoice = self._selected_invoices[conflict_id].currentData()
                payload["selected_invoice"] = invoice
            if conflict_id in self._selected_carriers:
                carrier = self._selected_carriers[conflict_id].currentData()
                payload["selected_carrier"] = carrier
            if conflict_id in self._selected_source_items:
                payload["selected_source_item_index"] = self._selected_source_items[
                    conflict_id
                ].currentData()
            resolutions[conflict_id] = payload
        return resolutions

    collect_resolutions = resolution_map

    def resolutions(self) -> list[dict[str, Any]]:
        return list(self.resolution_map().values())

    def _validate_and_accept(self) -> None:
        missing: list[str] = []
        for row, conflict in enumerate(self.conflicts):
            conflict_id = self._conflict_id(conflict, row)
            action = str(self._action_combos[conflict_id].currentData() or "")
            if action == "SELECT_ROW" and self._selected_rows.get(conflict_id) is None:
                missing.append(f"dòng {row + 1}: chưa chọn dòng BK")
            if action == "SELECT_FEE":
                combo = self._selected_fees.get(conflict_id)
                if combo is None or combo.currentData() in (None, ""):
                    missing.append(f"dòng {row + 1}: chưa chọn mã phí")
            if action == "SELECT_SHEET":
                combo = self._selected_sheets.get(conflict_id)
                if combo is None or combo.currentData() in (None, ""):
                    missing.append(f"dòng {row + 1}: chưa chọn sheet")
            if action == "SELECT_MONTH":
                combo = self._selected_months.get(conflict_id)
                if combo is None or combo.currentData() in (None, ""):
                    missing.append(f"dòng {row + 1}: chưa chọn tháng")
            if action == "SELECT_INVOICE":
                combo = self._selected_invoices.get(conflict_id)
                if combo is None or combo.currentData() in (None, ""):
                    missing.append(f"dòng {row + 1}: chưa chọn Số HĐ")
            if action == "SELECT_SOURCE_ITEM":
                combo = self._selected_source_items.get(conflict_id)
                if combo is None or combo.currentData() is None:
                    missing.append(f"dòng {row + 1}: chưa chọn dòng JSON")
            if action == "SELECT_CARRIER":
                combo = self._selected_carriers.get(conflict_id)
                if combo is None or combo.currentData() in (None, ""):
                    missing.append(f"dòng {row + 1}: chưa chọn bên vận tải")
            details = _value(conflict, "details", default={})
            if (
                action == "KEEP_EXISTING"
                and bool(
                    _value(
                        details,
                        "requires_existing_carrier_selection",
                        default=False,
                    )
                )
            ):
                combo = self._selected_carriers.get(conflict_id)
                if combo is None or combo.currentData() in (None, ""):
                    missing.append(
                        f"dòng {row + 1}: chọn mã vận tải hiệu lực đang có"
                    )
        if missing:
            self.validation_label.setText(" • ".join(missing))
            return
        self.validation_label.clear()
        self.accept()


class PaymentNewRowsDialog(QDialog):
    """Cho phép chọn các dòng BK mới sẽ được thêm vào file Thanh toán."""

    COLUMNS = (
        "Nhập",
        "Loại sheet",
        "Dòng BK",
        "SQT",
        "Container",
        "Các khoản sẽ ghi",
        "Trạng thái",
    )

    def __init__(
        self,
        items: Sequence[Any],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.items = list(items)
        self._checkbox_items: list[QTableWidgetItem] = []
        self.setObjectName("paymentNewRowsDialog")
        self.setWindowTitle("Quản lý dòng mới BK → Thanh toán")
        self.resize(980, 560)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        note = QLabel(
            f"Có {len(self.items)} dòng chưa tồn tại trong file Thanh toán. "
            "Mặc định chọn tất cả; dòng bỏ chọn sẽ xuất hiện lại ở lần đồng bộ sau."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        selection = QHBoxLayout()
        select_all = QPushButton("Chọn tất cả")
        select_none = QPushButton("Bỏ chọn tất cả")
        select_all.clicked.connect(
            lambda: self._set_all(Qt.CheckState.Checked)
        )
        select_none.clicked.connect(
            lambda: self._set_all(Qt.CheckState.Unchecked)
        )
        selection.addWidget(select_all)
        selection.addWidget(select_none)
        selection.addStretch()
        layout.addLayout(selection)

        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setObjectName("paymentNewRowsTable")
        self.table.setHorizontalHeaderLabels(list(self.COLUMNS))
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.table.horizontalHeader().setSectionResizeMode(
            5, QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self.table, 1)

        labels = {
            "sea_freight": "Cước biển",
            "north_freight": "Cước MB",
            "empty_lift": "Nâng vỏ",
            "loaded_drop": "Hạ hàng",
            "loaded_lift": "Nâng hàng",
            "empty_drop": "Hạ vỏ",
            "south_freight": "VTN",
            "storage": "Lưu cont",
            "overweight": "Quá tải",
            "vs_do": "VS + D/O",
            "command_fee": "Làm lệnh",
            "repair": "Sửa chữa",
        }
        for source in self.items:
            row = self.table.rowCount()
            self.table.insertRow(row)
            checkbox = QTableWidgetItem()
            checkbox.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsSelectable
            )
            checkbox.setCheckState(Qt.CheckState.Checked)
            checkbox.setData(
                Qt.ItemDataRole.UserRole,
                str(_value(source, "item_id", "id", default="")),
            )
            self._checkbox_items.append(checkbox)
            self.table.setItem(row, 0, checkbox)
            values = _value(source, "values", default={}) or {}
            invoices = _value(source, "invoice_values", default={}) or {}
            parts: list[str] = []
            for key, value in values.items():
                if value in (None, "", 0):
                    continue
                detail = f"{labels.get(key, key)}: {_format_amount(value)}"
                invoice = _value(invoices, key, default=None)
                if invoice not in (None, ""):
                    detail += f" – Số HĐ: {invoice}"
                parts.append(detail)
            amounts = "; ".join(parts)
            cells = (
                _value(source, "target_type", default="—"),
                _value(source, "source_row"),
                _value(source, "sqt"),
                _value(source, "container"),
                amounts or "Không có khoản tiền",
                "Hợp lệ",
            )
            for column, value in enumerate(cells, 1):
                item = QTableWidgetItem(_display(value))
                item.setToolTip(_display(value))
                self.table.setItem(row, column, item)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Xác nhận")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Hủy")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _set_all(self, state: Qt.CheckState) -> None:
        for item in self._checkbox_items:
            item.setCheckState(state)

    @property
    def selected_item_ids(self) -> list[str]:
        return [
            str(item.data(Qt.ItemDataRole.UserRole))
            for item in self._checkbox_items
            if item.checkState() == Qt.CheckState.Checked
        ]


ExcelConflictDialog = ConflictResolutionDialog
AggregateConflictDialog = ConflictResolutionDialog
TargetMonthDialog = MonthSelectionDialog


__all__ = [
    "AggregateConflictDialog",
    "ConflictResolutionDialog",
    "ExcelConflictDialog",
    "ManualRowPickerDialog",
    "MonthSelectionDialog",
    "PaymentNewRowsDialog",
    "RepostSelectionDialog",
    "TargetMonthDialog",
    "VALID_FEE_CODES",
]
