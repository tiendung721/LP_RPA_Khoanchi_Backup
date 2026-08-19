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


def _enable_user_sorting(table: QTableWidget) -> None:
    """Bật sort theo header nhưng giữ nguyên thứ tự nghiệp vụ ban đầu."""

    header = table.horizontalHeader()
    header.setSectionsClickable(True)
    header.setSortIndicatorShown(True)
    header.setSortIndicator(-1, Qt.SortOrder.AscendingOrder)
    table.setSortingEnabled(True)


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
        multi_select: bool = False,
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
        self.multi_select = multi_select
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
            QAbstractItemView.SelectionMode.ExtendedSelection
            if self.multi_select
            else QAbstractItemView.SelectionMode.SingleSelection
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
            bool(self.table.selectionModel().selectedRows())
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

    @property
    def selected_candidates(self) -> list[Any]:
        result: list[Any] = []
        for index in sorted(self.table.selectionModel().selectedRows(), key=lambda item: item.row()):
            item = self.table.item(index.row(), 0)
            if item is not None:
                result.append(item.data(Qt.ItemDataRole.UserRole))
        return result

    @property
    def selected_sheet_names(self) -> list[str]:
        return [name for candidate in self.selected_candidates if (name := _sheet_name(candidate))]

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


class PostingAllocationDialog(QDialog):
    """Phân bổ từng chứng từ/nhóm hóa đơn vào các sheet BK đã tồn tại."""

    def __init__(self, plan: Any, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("postingAllocationDialog")
        self.setWindowTitle("Phân bổ chứng từ vào sheet BK")
        self.resize(980, 520)
        self.groups = list(_sequence(_value(plan, "source_groups", default=())))
        candidates = list(_sequence(_value(plan, "sheet_candidates", default=())))
        self.sheet_names = [name for item in candidates if (name := _sheet_name(item))]
        self._combos: dict[str, QComboBox] = {}
        self._split_document_ids: set[str] = set()
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        note = QLabel(
            "Ngày hóa đơn chỉ dùng để gợi ý. Hãy chọn sheet đích cho mọi chứng từ; "
            "hồ sơ cước biển đã đối soát sẽ bị khóa sheet."
        )
        note.setWordWrap(True)
        layout.addWidget(note)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(
            ["Chứng từ / nhóm HĐ", "Số khoản", "Ngày HĐ", "Gợi ý", "Sheet đích"]
        )
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )
        for group in self.groups:
            row = self.table.rowCount()
            self.table.insertRow(row)
            group_id = str(_value(group, "group_id"))
            document = str(_value(group, "source_document_name", default=""))
            invoice = _value(group, "invoice_no")
            label = document + (f" — HĐ {invoice}" if invoice else "")
            self.table.setItem(row, 0, QTableWidgetItem(label))
            source_indices = list(_sequence(_value(group, "source_item_indices", default=())))
            self.table.setItem(row, 1, QTableWidgetItem(str(len(source_indices))))
            dates = list(_sequence(_value(group, "invoice_dates", default=())))
            self.table.setItem(row, 2, QTableWidgetItem(", ".join(map(str, dates)) or "—"))
            suggested = _value(group, "suggested_sheet")
            self.table.setItem(row, 3, QTableWidgetItem(_display(suggested)))
            combo = QComboBox()
            combo.addItem("— Chọn sheet —", "")
            for sheet_name in self.sheet_names:
                combo.addItem(sheet_name, sheet_name)
            target = _value(group, "target_sheet")
            if target:
                combo.setCurrentIndex(max(0, combo.findData(str(target))))
            elif suggested:
                combo.setCurrentIndex(max(0, combo.findData(str(suggested))))
            locked = bool(_value(group, "target_locked", default=False))
            combo.setEnabled(not locked)
            if locked:
                combo.setToolTip("Sheet đã khóa theo hồ sơ đối soát cước biển.")
            combo.currentIndexChanged.connect(self._update_summary)
            self.table.setCellWidget(row, 4, combo)
            self._combos[group_id] = combo
        layout.addWidget(self.table, 1)
        self.split_button = QPushButton("Tách tài liệu đã chọn theo hóa đơn")
        self.split_button.setToolTip(
            "Tạo nhóm riêng cho từng số hóa đơn trong cùng tài liệu trước khi phân bổ."
        )
        self.split_button.clicked.connect(self._request_invoice_split)
        self.table.itemSelectionChanged.connect(self._update_split_action)
        layout.addWidget(self.split_button)
        self.summary_label = QLabel()
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)
        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Phân tích")
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Hủy")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self._update_summary()
        if self.table.rowCount():
            self.table.selectRow(0)
        self._update_split_action()

    def _selected_group(self) -> Any | None:
        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            return None
        row = indexes[0].row()
        return self.groups[row] if 0 <= row < len(self.groups) else None

    def _update_split_action(self) -> None:
        group = self._selected_group()
        self.split_button.setEnabled(
            group is not None
            and bool(_value(group, "can_split_by_invoice", default=False))
        )

    def _request_invoice_split(self) -> None:
        group = self._selected_group()
        if group is None or not bool(
            _value(group, "can_split_by_invoice", default=False)
        ):
            return
        document_id = str(_value(group, "source_document_id", default=""))
        if document_id:
            self._split_document_ids.add(document_id)
            self.accept()

    def _update_summary(self, *_args: Any) -> None:
        totals: dict[str, tuple[int, int]] = {}
        for group in self.groups:
            group_id = str(_value(group, "group_id"))
            sheet = str(self._combos[group_id].currentData() or "")
            if not sheet:
                continue
            documents, items = totals.get(sheet, (0, 0))
            totals[sheet] = (
                documents + 1,
                items + len(_sequence(_value(group, "source_item_indices", default=()))),
            )
        self.summary_label.setText(
            " | ".join(
                f"{sheet}: {documents} nhóm / {items} khoản"
                for sheet, (documents, items) in totals.items()
            )
            or "Chưa phân bổ chứng từ."
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(
            bool(self.groups)
            and all(str(combo.currentData() or "") for combo in self._combos.values())
        )

    @property
    def group_target_sheets(self) -> dict[str, str]:
        return {
            group_id: str(combo.currentData())
            for group_id, combo in self._combos.items()
            if combo.currentData()
        }

    @property
    def split_document_ids(self) -> set[str]:
        return set(self._split_document_ids)


class DailySyncAllocationDialog(QDialog):
    """Ánh xạ riêng từng sheet Hàng ngày sang sheet BK đúng năm."""

    def __init__(self, plan: Any, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Ánh xạ các tháng Hàng ngày vào BK")
        self.resize(700, 430)
        self.plan = plan
        self._combos: dict[str, QComboBox] = {}
        layout = QVBoxLayout(self)
        note = QLabel("Chọn sheet BK đích cho từng sheet nguồn. Mỗi nguồn chỉ được xử lý một lần.")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Sheet Hàng ngày", "Sheet BK đích"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        mappings = dict(_value(plan, "source_target_sheets", default={}) or {})
        candidates = list(_sequence(_value(plan, "month_candidates", default=())))
        for source_sheet, target in mappings.items():
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(str(source_sheet)))
            combo = QComboBox()
            combo.addItem("— Chọn sheet —", "")
            for candidate in candidates:
                if str(_value(candidate, "source_sheet")) == str(source_sheet):
                    name = _sheet_name(candidate)
                    combo.addItem(name, name)
            if target:
                combo.setCurrentIndex(max(0, combo.findData(str(target))))
            combo.currentIndexChanged.connect(self._update_action)
            self.table.setCellWidget(row, 1, combo)
            self._combos[str(source_sheet)] = combo
        layout.addWidget(self.table, 1)
        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Phân tích")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self._update_action()

    def _update_action(self, *_args: Any) -> None:
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(
            bool(self._combos)
            and all(combo.currentData() for combo in self._combos.values())
        )

    @property
    def source_target_sheets(self) -> dict[str, str]:
        return {
            source: str(combo.currentData())
            for source, combo in self._combos.items()
            if combo.currentData()
        }


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

        self.table = QTableWidget(0, 8)
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
                "Bên vận tải",
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
                _value(candidate, "carrier", "transport_provider", "transport"),
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
        _enable_user_sorting(self.table)

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
    "CONTAINER_NOT_FOUND": ("SELECT_ROW", "SKIP"),
    "MULTIPLE_CONTAINER_MATCH": ("SELECT_ROW", "SKIP"),
    "REPEATED_SOURCE_CONTAINER": ("SELECT_ROW", "SKIP"),
    "TARGET_CELL_OCCUPIED": ("KEEP_EXISTING", "OVERWRITE"),
    "TARGET_CELL_FORMULA": ("KEEP_FORMULA", "OVERWRITE"),
    "TARGET_CELL_TEXT": ("KEEP_EXISTING", "OVERWRITE"),
    "UNKNOWN_FEE_CODE": ("SELECT_FEE", "SKIP"),
    "FEE_COLUMN_MISSING": ("SKIP", "CANCEL_ALL"),
    "BL_ONLY_NO_CONTAINER": ("SELECT_ROW", "SKIP"),
    "PARTIAL_KEY_MATCH": ("SELECT_ROW", "SKIP"),
    "NEGATIVE_ADJUSTMENT": ("ADD", "SKIP", "CANCEL_ALL"),
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
    "KEEP_FORMULA": "Giữ nguyên",
    "OVERWRITE": "Ghi đè",
    "ADD": "Áp dụng điều chỉnh giảm",
    "APPEND_CARRIER": "Ghi thêm",
    "POST_UNPOSTED_ONLY": "Chỉ nhập khoản chưa xử lý",
    "REANALYZE": "Đọc lại dữ liệu",
    "RETRY": "Thử lại",
}


BULK_ACTION_GROUPS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("Bỏ qua", ("SKIP", "SKIP_INVOICE", "SKIP_INVALID")),
    ("Giữ nguyên", ("KEEP_EXISTING", "KEEP_FORMULA")),
    ("Ghi đè", ("OVERWRITE",)),
)


ROW_SELECTION_CONFLICT_TYPES = frozenset(
    {
        "CONTAINER_NOT_FOUND",
        "MULTIPLE_CONTAINER_MATCH",
        "REPEATED_SOURCE_CONTAINER",
        "BL_ONLY_NO_CONTAINER",
        "PARTIAL_KEY_MATCH",
    }
)

TARGET_VALUE_CONFLICT_TYPES = frozenset(
    {
        "TARGET_CELL_OCCUPIED",
        "TARGET_CELL_FORMULA",
        "TARGET_CELL_TEXT",
    }
)

EXPLICIT_ACTION_CONFLICT_TYPES = (
    ROW_SELECTION_CONFLICT_TYPES | TARGET_VALUE_CONFLICT_TYPES
)


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

        _enable_user_sorting(self.table)

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
        "Bên vận tải",
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
        initial_resolutions: Mapping[str, Any] | None = None,
        restore_info: Mapping[str, Any] | None = None,
        issues: Sequence[Any] | None = None,
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
        self.initial_resolutions = dict(initial_resolutions or {})
        self.restore_info = dict(restore_info or {})
        self.issues = list(issues or ())
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
        self._conflicts_by_id: dict[str, Any] = {}
        self._inactive_conflict_ids: set[str] = set()
        self.setObjectName("excelConflictResolutionDialog")
        self.setWindowTitle("Xử lý xung đột Excel")
        self.resize(1320, 660)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        self.title_label = QLabel(
            f"Có {len(self.conflicts)} mục cần xử lý trước khi ghi workbook."
        )
        self.title_label.setStyleSheet("font-size: 12pt; font-weight: 700;")
        layout.addWidget(self.title_label)
        note = QLabel(
            "Kiểm tra giá trị hiện tại và chọn cách xử lý cho từng mục. "
            "Workbook chỉ được ghi sau khi toàn bộ lựa chọn hợp lệ."
        )
        note.setWordWrap(True)
        note.setProperty("muted", True)
        layout.addWidget(note)

        if bool(self.restore_info.get("found")):
            status_labels = {
                "ACTIVE": "Đang xử lý",
                "SUCCEEDED": "Thành công",
                "FAILED": "Thất bại",
                "CANCELLED": "Đã hủy",
            }
            updated_at = str(self.restore_info.get("updated_at") or "")
            display_time = updated_at.replace("T", " ").split("+")[0]
            status = status_labels.get(
                str(self.restore_info.get("status") or ""),
                str(self.restore_info.get("status") or "Không xác định"),
            )
            restored_here = sum(
                self._conflict_id(conflict, row) in self.initial_resolutions
                for row, conflict in enumerate(self.conflicts)
            )
            if bool(self.restore_info.get("target_compatible", True)):
                message = (
                    f"Đã khôi phục {restored_here}/{len(self.conflicts)} lựa chọn "
                    f"từ lần xử lý lúc {display_time} (kết quả: {status})."
                )
            else:
                message = (
                    f"Đã tìm thấy lần xử lý lúc {display_time} "
                    f"(kết quả: {status}), nhưng workbook đích đã thay đổi; "
                    "cần chọn lại toàn bộ."
                )
            self.restore_label = QLabel(message)
            self.restore_label.setObjectName("conflictRestoreStatusLabel")
            self.restore_label.setWordWrap(True)
            self.restore_label.setProperty("status", "info")
            layout.addWidget(self.restore_label)

        bulk_layout = QHBoxLayout()
        bulk_label = QLabel("Xử lý hàng loạt:")
        bulk_label.setObjectName("conflictBulkActionLabel")
        self.bulk_action_combo = QComboBox()
        self.bulk_action_combo.setObjectName("conflictBulkActionCombo")
        self.bulk_action_combo.setMinimumWidth(230)
        self.bulk_apply_button = QPushButton("Áp dụng cho toàn bộ")
        self.bulk_apply_button.setObjectName("conflictBulkApplyButton")
        self.bulk_apply_button.setEnabled(False)
        self.bulk_result_label = QLabel()
        self.bulk_result_label.setObjectName("conflictBulkResultLabel")
        self.bulk_result_label.setProperty("muted", True)
        bulk_layout.addWidget(bulk_label)
        bulk_layout.addWidget(self.bulk_action_combo)
        bulk_layout.addWidget(self.bulk_apply_button)
        bulk_layout.addWidget(self.bulk_result_label, 1)
        layout.addLayout(bulk_layout)

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
            11, QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self.table, 1)

        for index, conflict in enumerate(self.conflicts):
            self._add_conflict(index, conflict)
        self._wire_action_dependencies()
        self._refresh_bulk_actions()
        _enable_user_sorting(self.table)

        self.bulk_action_combo.currentIndexChanged.connect(
            self._update_bulk_apply_enabled
        )
        self.bulk_apply_button.clicked.connect(self._apply_bulk_action)

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
            "Xác nhận lựa chọn"
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Hủy")
        self.buttons.accepted.connect(self._validate_and_accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self.show_issues(self.issues)

    def set_review(
        self,
        conflicts: Sequence[Any],
        *,
        issues: Sequence[Any] = (),
        initial_resolutions: Mapping[str, Any] | None = None,
    ) -> None:
        """Reuse this dialog for the next correction round."""

        carried = self.resolution_map() if self._action_combos else {}
        carried.update(dict(initial_resolutions or {}))
        self.conflicts = list(conflicts)
        self.initial_resolutions = carried
        self.issues = list(issues)
        self.table.setSortingEnabled(False)
        self.table.clearContents()
        self.table.setRowCount(0)
        self._action_combos.clear()
        self._selected_rows.clear()
        self._selected_source_sheets.clear()
        self._selected_source_items.clear()
        self._selected_fees.clear()
        self._selected_sheets.clear()
        self._selected_months.clear()
        self._selected_invoices.clear()
        self._selected_carriers.clear()
        self._selector_buttons.clear()
        self._conflicts_by_id.clear()
        self._inactive_conflict_ids.clear()
        for index, conflict in enumerate(self.conflicts):
            self._add_conflict(index, conflict)
        self._wire_action_dependencies()
        self._refresh_bulk_actions()
        self.bulk_result_label.clear()
        self.table.setSortingEnabled(True)
        self.setResult(0)
        self.show_issues(self.issues)

    @staticmethod
    def _issue_conflict_ids(issue: Any) -> tuple[str, ...]:
        raw = _value(issue, "conflict_ids", default=())
        values = tuple(str(value) for value in _sequence(raw) if str(value))
        if values:
            return values
        single = _value(issue, "conflict_id", default=None)
        return (str(single),) if single not in (None, "") else ()

    def show_issues(self, issues: Sequence[Any]) -> None:
        """Paint latest validation issues without reacting to later edits."""

        self.issues = list(issues)
        messages: dict[str, list[str]] = {}
        for issue in self.issues:
            message = str(_value(issue, "message", default="Chưa hợp lệ."))
            for conflict_id in self._issue_conflict_ids(issue):
                messages.setdefault(conflict_id, []).append(message)

        problem_widget_style = (
            "QComboBox, QPushButton { background-color: #fff7d6; "
            "border: 1px solid #d6a700; }"
        )
        first_problem_row: int | None = None
        problem_column = self.COLUMNS.index("Vấn đề")
        for row in range(self.table.rowCount()):
            identity_item = self.table.item(row, 0)
            conflict_id = (
                str(identity_item.data(Qt.ItemDataRole.UserRole + 1))
                if identity_item is not None
                else ""
            )
            row_messages = messages.get(conflict_id, [])
            for column in range(self.table.columnCount()):
                widget = self.table.cellWidget(row, column)
                if widget is not None:
                    widget.setStyleSheet(
                        problem_widget_style if row_messages else ""
                    )
            if row_messages:
                if first_problem_row is None:
                    first_problem_row = row
                issue_item = self.table.item(row, problem_column)
                if issue_item is not None:
                    original = issue_item.toolTip() or issue_item.text()
                    latest = " • ".join(dict.fromkeys(row_messages))
                    issue_item.setText(latest)
                    issue_item.setToolTip(
                        f"{latest}\n\nXung đột ban đầu: {original}"
                    )

        if messages:
            count = len(messages)
            self.title_label.setText(
                f"Có {count} dòng cần sửa trong {len(self.conflicts)} mục xung đột."
            )
            self.validation_label.setText(
                "Kiểm tra các ô được đánh dấu và chọn đủ cách xử lý."
            )
        else:
            self.title_label.setText(
                f"Có {len(self.conflicts)} mục cần xử lý trước khi ghi workbook."
            )
            self.validation_label.clear()
        if first_problem_row is not None:
            self.table.selectRow(first_problem_row)
            item = self.table.item(first_problem_row, 0)
            if item is not None:
                self.table.scrollToItem(
                    item, QAbstractItemView.ScrollHint.PositionAtCenter
                )

    def _refresh_bulk_actions(self) -> None:
        """Liệt kê các cách xử lý có thể áp dụng cho ít nhất một dòng."""

        available_codes: list[str] = []
        for combo in self._action_combos.values():
            for index in range(combo.count()):
                code = str(combo.itemData(index) or "")
                if code and code not in available_codes:
                    available_codes.append(code)

        grouped_codes: set[str] = set()
        options: list[tuple[str, tuple[str, ...], int]] = []
        for label, action_codes in BULK_ACTION_GROUPS:
            present_codes = tuple(
                code for code in action_codes if code in available_codes
            )
            if not present_codes:
                continue
            grouped_codes.update(present_codes)
            applicable_count = sum(
                any(combo.findData(code) >= 0 for code in present_codes)
                for combo in self._action_combos.values()
            )
            options.append((label, present_codes, applicable_count))

        for code in available_codes:
            if code in grouped_codes:
                continue
            applicable_count = sum(
                combo.findData(code) >= 0
                for combo in self._action_combos.values()
            )
            options.append(
                (ACTION_LABELS.get(code, code), (code,), applicable_count)
            )

        was_blocked = self.bulk_action_combo.blockSignals(True)
        try:
            self.bulk_action_combo.clear()
            self.bulk_action_combo.addItem("— Chọn cách xử lý —", ())
            for label, action_codes, applicable_count in options:
                self.bulk_action_combo.addItem(
                    f"{label} ({applicable_count} dòng)", action_codes
                )
            self.bulk_action_combo.setCurrentIndex(0)
        finally:
            self.bulk_action_combo.blockSignals(was_blocked)
        self._update_bulk_apply_enabled()

    def _update_bulk_apply_enabled(self, _index: int = -1) -> None:
        action_codes = self.bulk_action_combo.currentData()
        self.bulk_apply_button.setEnabled(bool(action_codes))

    def _apply_bulk_action(self) -> None:
        """Đặt lựa chọn chung; từng dòng vẫn có thể được sửa lại sau đó."""

        raw_codes = self.bulk_action_combo.currentData()
        action_codes = tuple(str(code) for code in _sequence(raw_codes) if code)
        if not action_codes:
            return

        applied_count = 0
        for conflict_id, combo in self._action_combos.items():
            if conflict_id in self._inactive_conflict_ids:
                continue
            target_index = next(
                (
                    combo.findData(code)
                    for code in action_codes
                    if combo.findData(code) >= 0
                ),
                -1,
            )
            if target_index < 0:
                continue
            combo.setCurrentIndex(target_index)
            applied_count += 1

        total_count = len(self._action_combos)
        selected_label = self.bulk_action_combo.currentText().rsplit(" (", 1)[0]
        if applied_count == total_count:
            message = (
                f"Đã áp dụng “{selected_label}” cho toàn bộ "
                f"{total_count} dòng."
            )
        else:
            remaining_count = total_count - applied_count
            message = (
                f"Đã áp dụng “{selected_label}” cho {applied_count}/{total_count} "
                f"dòng; {remaining_count} dòng không có lựa chọn này."
            )
        self.bulk_result_label.setText(message)

    def _add_conflict(self, row: int, conflict: Any) -> None:
        self.table.insertRow(row)
        conflict_id = self._conflict_id(conflict, row)
        self._conflicts_by_id[conflict_id] = conflict
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
        carrier_value = _value(conflict, "carrier", "carrier_value")
        if carrier_value in (None, ""):
            carrier_value = _value(
                details,
                "carrier_incoming",
                "incoming_carrier",
                default=None,
            )
        carrier_display = (
            str(carrier_value)
            if carrier_value not in (None, "")
            else ", ".join(str(value) for value in carrier_candidates)
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
            carrier_display or None,
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
        if conflict_type in EXPLICIT_ACTION_CONFLICT_TYPES:
            action_combo.addItem("", "")
        elif conflict_type == "NEGATIVE_ADJUSTMENT":
            action_combo.addItem("— Chọn cách xử lý —", "")
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
        action_combo.setCurrentIndex(
            0
            if conflict_type in EXPLICIT_ACTION_CONFLICT_TYPES
            or conflict_type == "NEGATIVE_ADJUSTMENT"
            else max(0, default_index)
        )
        self._action_combos[conflict_id] = action_combo
        self.table.setCellWidget(row, 12, action_combo)

        selector = self._selector_for(row, conflict, conflict_id, conflict_type)
        if selector is not None:
            self.table.setCellWidget(row, 13, selector)
        else:
            self.table.setItem(row, 13, QTableWidgetItem("—"))
        self._apply_initial_resolution(conflict_id, action_combo)

    def _apply_initial_resolution(
        self, conflict_id: str, action_combo: QComboBox
    ) -> None:
        """Điền lựa chọn gần nhất vào bảng để user vẫn có thể sửa."""

        raw = self.initial_resolutions.get(conflict_id)
        if not isinstance(raw, Mapping):
            return

        action = _code(raw.get("action"))
        action_index = action_combo.findData(action)
        if action_index >= 0:
            action_combo.setCurrentIndex(action_index)

        if conflict_id in self._selected_rows:
            selected_row = raw.get("selected_row", raw.get("row"))
            if selected_row is not None:
                self._selected_rows[conflict_id] = selected_row
                selected_sheet = raw.get(
                    "selected_source_sheet", raw.get("selected_sheet")
                )
                if selected_sheet not in (None, ""):
                    self._selected_source_sheets[conflict_id] = str(selected_sheet)
                button = self._selector_buttons.get(conflict_id)
                if button is not None:
                    label = f"Dòng {selected_row}"
                    if selected_sheet not in (None, ""):
                        label = f"{selected_sheet} – dòng {selected_row}"
                    button.setText(label)

        combo_values = (
            (self._selected_fees, ("selected_fee", "fee")),
            (
                self._selected_sheets,
                ("selected_sheet_name", "selected_sheet", "sheet_name"),
            ),
            (self._selected_months, ("selected_month", "month")),
            (self._selected_invoices, ("selected_invoice",)),
            (self._selected_carriers, ("selected_carrier",)),
            (self._selected_source_items, ("selected_source_item_index",)),
        )
        for widgets, names in combo_values:
            combo = widgets.get(conflict_id)
            if combo is None:
                continue
            value = next((raw.get(name) for name in names if name in raw), None)
            if value is None:
                continue
            index = combo.findData(value)
            if index >= 0:
                combo.setCurrentIndex(index)

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
            "NEGATIVE_ADJUSTMENT",
        }:
            button = QPushButton("Chọn dòng…")
            button.setObjectName(f"selectConflictRow_{row}")
            button.clicked.connect(
                lambda _checked=False, c=conflict, key=conflict_id, b=button: (
                    self._pick_row(c, key, b)
                )
            )
            self._selected_rows[conflict_id] = None
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
        if conflict_type in TARGET_VALUE_CONFLICT_TYPES:
            actions = tuple(action for action in actions if _code(action) != "SKIP")
        if conflict_type in ROW_SELECTION_CONFLICT_TYPES:
            actions = tuple(
                action for action in actions if _code(action) != "CANCEL_ALL"
            )
            order = {"SELECT_ROW": 0, "SKIP": 1}
            actions = tuple(
                sorted(actions, key=lambda action: order.get(_code(action), 99))
            )
        elif conflict_type in TARGET_VALUE_CONFLICT_TYPES:
            order = {
                "KEEP_EXISTING": 0,
                "KEEP_FORMULA": 0,
                "OVERWRITE": 1,
            }
            actions = tuple(
                sorted(actions, key=lambda action: order.get(_code(action), 99))
            )
        return actions

    @staticmethod
    def _dependency_key(conflict: Any) -> tuple[Any, ...]:
        item_index = _value(conflict, "item_index", default=None)
        if item_index is not None:
            return ("item", int(item_index))
        source_key = (
            "source",
            _value(conflict, "container", "container_number"),
            _value(conflict, "bl", "bill_of_lading"),
            _value(conflict, "selected_fee", "fee", "fee_code"),
            _value(conflict, "amount", "incoming_amount"),
        )
        if all(value in (None, "") for value in source_key[1:]):
            return ("object", id(conflict))
        return source_key

    def _wire_action_dependencies(self) -> None:
        parents: dict[tuple[Any, ...], str] = {}
        children: dict[tuple[Any, ...], list[str]] = {}
        for conflict_id, conflict in self._conflicts_by_id.items():
            conflict_type = _code(
                _value(conflict, "conflict_type", "type", "kind")
            )
            key = self._dependency_key(conflict)
            if conflict_type in ROW_SELECTION_CONFLICT_TYPES:
                parents[key] = conflict_id
            elif conflict_type in TARGET_VALUE_CONFLICT_TYPES:
                children.setdefault(key, []).append(conflict_id)

        for key, parent_id in parents.items():
            child_ids = tuple(children.get(key, ()))
            combo = self._action_combos[parent_id]
            combo.currentIndexChanged.connect(
                lambda _index, current=parent_id, linked=child_ids: (
                    self._sync_action_dependency(current, linked)
                )
            )
            self._sync_action_dependency(parent_id, child_ids)

    def _sync_action_dependency(
        self, parent_id: str, child_ids: Sequence[str]
    ) -> None:
        parent_action = str(
            self._action_combos[parent_id].currentData() or ""
        )
        active = parent_action == "SELECT_ROW"
        if not active:
            self._selected_rows[parent_id] = None
            self._selected_source_sheets.pop(parent_id, None)
        button = self._selector_buttons.get(parent_id)
        if button is not None:
            button.setEnabled(active)
            if not active:
                button.setText("Chọn dòng…")

        for child_id in child_ids:
            combo = self._action_combos[child_id]
            combo.setEnabled(active)
            if active:
                self._inactive_conflict_ids.discard(child_id)
                continue
            self._inactive_conflict_ids.add(child_id)
            blank_index = combo.findData("")
            if blank_index >= 0:
                combo.setCurrentIndex(blank_index)

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
        missing_issues: list[dict[str, Any]] = []

        def add_missing(conflict_id: str, text: str, message: str) -> None:
            missing.append(text)
            missing_issues.append(
                {"conflict_ids": (conflict_id,), "message": message}
            )

        for row, conflict in enumerate(self.conflicts):
            conflict_id = self._conflict_id(conflict, row)
            if conflict_id in self._inactive_conflict_ids:
                continue
            action = str(self._action_combos[conflict_id].currentData() or "")
            if not action:
                add_missing(
                    conflict_id,
                    f"dòng {row + 1}: chưa chọn cách xử lý",
                    "Chưa chọn cách xử lý.",
                )
            if action == "SELECT_ROW" and self._selected_rows.get(conflict_id) is None:
                add_missing(
                    conflict_id,
                    f"dòng {row + 1}: chưa chọn dòng BK",
                    "Chưa chọn dòng BK.",
                )
            if action == "SELECT_FEE":
                combo = self._selected_fees.get(conflict_id)
                if combo is None or combo.currentData() in (None, ""):
                    add_missing(
                        conflict_id,
                        f"dòng {row + 1}: chưa chọn mã phí",
                        "Chưa chọn mã phí.",
                    )
            if action == "SELECT_SHEET":
                combo = self._selected_sheets.get(conflict_id)
                if combo is None or combo.currentData() in (None, ""):
                    add_missing(
                        conflict_id,
                        f"dòng {row + 1}: chưa chọn sheet",
                        "Chưa chọn sheet.",
                    )
            if action == "SELECT_MONTH":
                combo = self._selected_months.get(conflict_id)
                if combo is None or combo.currentData() in (None, ""):
                    add_missing(
                        conflict_id,
                        f"dòng {row + 1}: chưa chọn tháng",
                        "Chưa chọn tháng.",
                    )
            if action == "SELECT_INVOICE":
                combo = self._selected_invoices.get(conflict_id)
                if combo is None or combo.currentData() in (None, ""):
                    add_missing(
                        conflict_id,
                        f"dòng {row + 1}: chưa chọn Số HĐ",
                        "Chưa chọn Số HĐ.",
                    )
            if action == "SELECT_SOURCE_ITEM":
                combo = self._selected_source_items.get(conflict_id)
                if combo is None or combo.currentData() is None:
                    add_missing(
                        conflict_id,
                        f"dòng {row + 1}: chưa chọn dòng JSON",
                        "Chưa chọn dòng JSON.",
                    )
            if action == "SELECT_CARRIER":
                combo = self._selected_carriers.get(conflict_id)
                if combo is None or combo.currentData() in (None, ""):
                    add_missing(
                        conflict_id,
                        f"dòng {row + 1}: chưa chọn bên vận tải",
                        "Chưa chọn bên vận tải.",
                    )
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
                    add_missing(
                        conflict_id,
                        f"dòng {row + 1}: chọn mã vận tải hiệu lực đang có",
                        "Chưa chọn mã vận tải hiệu lực đang có.",
                    )
        if missing:
            self.show_issues(missing_issues)
            self.validation_label.setText(" • ".join(missing))
            return
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

        _enable_user_sorting(self.table)

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


class ExcelOutcomeDialog(QDialog):
    """Chi tiết phẳng được dựng từ execution manifest thực tế."""

    _FILTERS = (
        ("Tất cả", None),
        ("Đã ghi", {"WRITTEN", "PARTIAL"}),
        ("Không đổi", {"UNCHANGED"}),
        ("Giữ / bỏ qua", {"USER_KEPT", "USER_SKIPPED"}),
        ("Lỗi", {"INVALID_SOURCE", "FAILED"}),
    )

    def __init__(
        self,
        outcomes: Sequence[Any],
        *,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Chi tiết kết quả xử lý Excel")
        self.resize(1320, 680)
        self._rows: list[tuple[str, tuple[Any, ...]]] = []
        for outcome in outcomes:
            item_status = _code(_value(outcome, "status", default="UNCHANGED"))
            fields = _sequence(_value(outcome, "fields", default=())) or (None,)
            for field in fields:
                field_status = _code(
                    _value(field, "status", default=item_status)
                )
                target_sheet = _value(
                    field,
                    "target_sheet",
                    default=_value(outcome, "target_sheet"),
                )
                target_cell = _value(field, "target_cell")
                target_row = _value(
                    field,
                    "target_row",
                    default=_value(outcome, "target_row"),
                )
                destination = target_cell or (
                    f"dòng {target_row}" if target_row is not None else "—"
                )
                self._rows.append(
                    (
                        field_status,
                        (
                            _value(outcome, "container"),
                            _value(outcome, "sqt"),
                            _value(outcome, "bl"),
                            _value(field, "field_name", default="Dòng dữ liệu"),
                            _value(field, "source_value"),
                            target_sheet,
                            destination,
                            _value(field, "target_value_before"),
                            _value(field, "target_value_after"),
                            field_status,
                            _value(field, "reason"),
                        ),
                    )
                )

        layout = QVBoxLayout(self)
        intro = QLabel(
            "Mỗi dòng dưới đây là một thành phần đã được xét trong lần thực thi."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Lọc kết quả:"))
        self.filter_combo = QComboBox()
        for label, statuses in self._FILTERS:
            self.filter_combo.addItem(label, statuses)
        self.filter_combo.currentIndexChanged.connect(self._refresh)
        filter_row.addWidget(self.filter_combo)
        filter_row.addStretch(1)
        layout.addLayout(filter_row)

        self.table = QTableWidget(0, 11)
        self.table.setHorizontalHeaderLabels(
            [
                "Container",
                "SQT",
                "BL",
                "Loại dữ liệu",
                "Nguồn",
                "Sheet đích",
                "Ô / dòng đích",
                "Giá trị trước",
                "Giá trị sau",
                "Kết quả",
                "Lý do",
            ]
        )
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(10, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self._refresh()

    def _refresh(self, _index: int = 0) -> None:
        allowed = self.filter_combo.currentData()
        visible = [
            cells
            for status, cells in self._rows
            if allowed is None or status in allowed
        ]
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(visible))
        for row, cells in enumerate(visible):
            for column, value in enumerate(cells):
                item = QTableWidgetItem(_display(value))
                item.setToolTip(_display(value))
                self.table.setItem(row, column, item)
        _enable_user_sorting(self.table)


ExcelConflictDialog = ConflictResolutionDialog
AggregateConflictDialog = ConflictResolutionDialog
TargetMonthDialog = MonthSelectionDialog


__all__ = [
    "AggregateConflictDialog",
    "ConflictResolutionDialog",
    "DailySyncAllocationDialog",
    "ExcelOutcomeDialog",
    "ExcelConflictDialog",
    "ManualRowPickerDialog",
    "MonthSelectionDialog",
    "PostingAllocationDialog",
    "PaymentNewRowsDialog",
    "RepostSelectionDialog",
    "TargetMonthDialog",
    "VALID_FEE_CODES",
]
