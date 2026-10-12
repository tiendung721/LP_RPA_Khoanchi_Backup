"""Chi tiết theo số QT trong hộp xác nhận nhập khoản chi vào BK."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)


@dataclass(frozen=True, slots=True)
class PostingDetailRow:
    sheet_name: str
    sqt: int | None
    sqt_origin: str
    container: str | None
    fee: str
    amount: Any
    value_before: Any
    value_after: Any
    target_row: int | None
    target_cell: str | None
    result: str
    reason: str
    invoice_no: str | None
    source_document_name: str | None
    source_row: int | None
    source_count: int = 1

    @property
    def qt_label(self) -> str:
        if self.sqt is None:
            return "Chưa xác định QT"
        return f"QT {self.sqt}" if self.sqt_origin == "BK" else f"QT nguồn {self.sqt}"


@dataclass(frozen=True, slots=True)
class PostingDetailData:
    cells: tuple[PostingDetailRow, ...] = ()
    write: tuple[PostingDetailRow, ...] = ()
    already: tuple[PostingDetailRow, ...] = ()
    skip: tuple[PostingDetailRow, ...] = ()

    def rows_for(self, metric_index: int) -> tuple[PostingDetailRow, ...]:
        return (self.cells, self.write, self.already, self.skip)[metric_index]


_METRIC_TITLES = (
    "Ô BK thay đổi",
    "Khoản sẽ ghi",
    "Đã đúng sẵn",
    "Không ghi",
)
_METRIC_DESCRIPTIONS = (
    "Mỗi dòng là một ô phí BK dự kiến thay đổi. Khoản gộp cùng ô chỉ xuất hiện một lần.",
    "Mỗi dòng là một khoản phí nguồn dự kiến ghi vào BK.",
    "Số tiền khoản phí đã bằng số tiền trong BK; không ghi lại ô tiền.",
    "Khoản phí được giữ nguyên hoặc bỏ qua, kèm lý do không ghi.",
)
_RESULT_LABELS = {
    "write_new": "Ghi mới",
    "overwrite": "Ghi đè",
    "already": "Đã đúng sẵn",
    "keep": "Giữ nguyên",
    "skip": "Bỏ qua",
}


def _money(value: Any) -> str:
    if value in (None, ""):
        return "—"
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return f"{value:,.0f}".replace(",", ".") + " đ"
    return str(value)


def _text(value: Any) -> str:
    return "—" if value in (None, "") else str(value)


class PostingConfirmationDetails(QFrame):
    """Vùng chi tiết đọc thuần túy, giữ nguyên kế hoạch và quyết định ghi."""

    def __init__(self, data: PostingDetailData, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.data = data
        self.metric_index = 0
        self.setObjectName("postingConfirmationDetails")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(10)

        heading = QHBoxLayout()
        title_column = QVBoxLayout()
        title_column.setSpacing(2)
        self.title = QLabel()
        self.title.setObjectName("postingDetailsTitle")
        self.description = QLabel()
        self.description.setWordWrap(True)
        self.description.setObjectName("postingDetailsDescription")
        title_column.addWidget(self.title)
        title_column.addWidget(self.description)
        heading.addLayout(title_column, 1)
        self.toggle = QToolButton()
        self.toggle.setObjectName("postingDetailsToggle")
        self.toggle.setText("Thu gọn ▴")
        self.toggle.setCheckable(True)
        self.toggle.setChecked(True)
        self.toggle.toggled.connect(self._toggle_content)
        heading.addWidget(self.toggle)
        layout.addLayout(heading)

        self.content = QWidget()
        content_layout = QVBoxLayout(self.content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(9)
        filters = QHBoxLayout()
        search_label = QLabel("Tìm số QT hoặc container")
        self.search = QLineEdit()
        self.search.setObjectName("postingDetailsSearch")
        self.search.setPlaceholderText("Nhập số QT hoặc container...")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self._render)
        search_column = QVBoxLayout()
        search_column.setSpacing(3)
        search_column.addWidget(search_label)
        search_column.addWidget(self.search)
        filters.addLayout(search_column, 1)
        sheet_column = QVBoxLayout()
        sheet_column.setSpacing(3)
        sheet_column.addWidget(QLabel("Sheet BK"))
        self.sheet_filter = QComboBox()
        self.sheet_filter.setObjectName("postingDetailsSheetFilter")
        self.sheet_filter.addItem("Tất cả sheet", "")
        sheets = sorted(
            {row.sheet_name for rows in (data.cells, data.write, data.already, data.skip)
             for row in rows if row.sheet_name}
        )
        for sheet in sheets:
            self.sheet_filter.addItem(sheet, sheet)
        self.sheet_filter.currentIndexChanged.connect(self._render)
        sheet_column.addWidget(self.sheet_filter)
        filters.addLayout(sheet_column)
        self.count_label = QLabel()
        self.count_label.setObjectName("postingDetailsCount")
        filters.addWidget(self.count_label, 0, Qt.AlignmentFlag.AlignBottom)
        content_layout.addLayout(filters)

        self.tree = QTreeWidget()
        self.tree.setObjectName("postingDetailsTree")
        self.tree.setHeaderLabels(
            ("Số QT / Container", "Khoản phí", "Tiền nguồn", "BK hiện có", "Ô BK", "Kết quả")
        )
        self.tree.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.tree.setAlternatingRowColors(True)
        self.tree.setUniformRowHeights(True)
        self.tree.setMinimumHeight(210)
        self.tree.setMaximumHeight(340)
        self.tree.header().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tree.currentItemChanged.connect(self._show_selected)
        content_layout.addWidget(self.tree)

        self.detail_card = QFrame()
        self.detail_card.setObjectName("postingDetailsSelected")
        detail_layout = QVBoxLayout(self.detail_card)
        detail_layout.setContentsMargins(12, 10, 12, 10)
        detail_layout.setSpacing(7)
        self.selected_title = QLabel("Chọn một khoản phí để xem chi tiết.")
        self.selected_title.setObjectName("postingDetailsSelectedTitle")
        detail_layout.addWidget(self.selected_title)
        grid = QGridLayout()
        grid.setHorizontalSpacing(18)
        grid.setVerticalSpacing(5)
        self.fields: dict[str, QLabel] = {}
        field_specs = (
            ("container", "Container"),
            ("invoice", "Số HĐ nguồn"),
            ("before", "BK hiện có"),
            ("after", "Sau xác nhận"),
            ("location", "Vị trí BK"),
            ("source", "File nguồn"),
            ("reason", "Diễn giải"),
        )
        for index, (key, caption) in enumerate(field_specs):
            column = index % 2
            row = index // 2
            name = QLabel(caption + ":")
            name.setObjectName("postingDetailsFieldName")
            value = QLabel("—")
            value.setObjectName("postingDetailsFieldValue")
            value.setWordWrap(True)
            value.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            grid.addWidget(name, row, column * 2)
            grid.addWidget(value, row, column * 2 + 1)
            self.fields[key] = value
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)
        detail_layout.addLayout(grid)
        content_layout.addWidget(self.detail_card)
        layout.addWidget(self.content)
        self.set_metric(0)

    def _toggle_content(self, expanded: bool) -> None:
        self.content.setVisible(expanded)
        self.toggle.setText("Thu gọn ▴" if expanded else "Mở chi tiết ▾")

    def set_metric(self, metric_index: int) -> None:
        self.metric_index = metric_index
        self.title.setText("CHI TIẾT: " + _METRIC_TITLES[metric_index].upper())
        self.description.setText(_METRIC_DESCRIPTIONS[metric_index])
        self.toggle.setChecked(True)
        self._render()

    def set_sheet(self, sheet_name: str) -> None:
        index = self.sheet_filter.findData(sheet_name)
        if index >= 0:
            self.sheet_filter.setCurrentIndex(index)
        self.toggle.setChecked(True)

    def _render(self, *_args: Any) -> None:
        query = self.search.text().strip().casefold()
        selected_sheet = str(self.sheet_filter.currentData() or "")
        rows = [
            row for row in self.data.rows_for(self.metric_index)
            if (not selected_sheet or row.sheet_name == selected_sheet)
            and (not query or query in str(row.sqt or "").casefold()
                 or query in str(row.container or "").casefold())
        ]
        groups: dict[tuple[str, int | None, str], list[PostingDetailRow]] = defaultdict(list)
        for row in rows:
            groups[(row.sheet_name, row.sqt, row.sqt_origin)].append(row)
        unit = "ô" if self.metric_index == 0 else "khoản"
        self.count_label.setText(f"{len(rows)} {unit} · {len(groups)} nhóm QT")

        self.tree.blockSignals(True)
        self.tree.clear()
        first_child: QTreeWidgetItem | None = None
        for group_index, ((sheet, _sqt, _origin), members) in enumerate(groups.items()):
            result_counts = defaultdict(int)
            for member in members:
                result_counts[member.result] += 1
            statuses = " · ".join(
                f"{count} {_RESULT_LABELS[key].lower()}"
                for key in ("write_new", "overwrite", "already", "keep", "skip")
                if (count := result_counts[key])
            )
            group = QTreeWidgetItem(
                self.tree, (f"{members[0].qt_label} · {sheet}", statuses, "", "", "", "")
            )
            group.setFlags(group.flags() & ~Qt.ItemFlag.ItemIsSelectable)
            for row in members:
                child = QTreeWidgetItem(
                    group,
                    (
                        _text(row.container),
                        row.fee,
                        _money(row.amount),
                        _money(row.value_before),
                        _text(row.target_cell),
                        _RESULT_LABELS[row.result],
                    ),
                )
                child.setData(0, Qt.ItemDataRole.UserRole, row)
                if first_child is None:
                    first_child = child
            group.setExpanded(group_index == 0)
        self.tree.blockSignals(False)
        if first_child is not None:
            self.tree.setCurrentItem(first_child)
        else:
            self._show_selected(None)

    def _show_selected(self, current: QTreeWidgetItem | None, *_args: Any) -> None:
        row = current.data(0, Qt.ItemDataRole.UserRole) if current is not None else None
        if not isinstance(row, PostingDetailRow):
            self.selected_title.setText(
                "Không có khoản nào không ghi trong lượt này."
                if self.metric_index == 3 and not self.search.text().strip()
                else "Không tìm thấy khoản phù hợp."
            )
            for value in self.fields.values():
                value.setText("—")
            return
        self.selected_title.setText(
            f"{_RESULT_LABELS[row.result]}: {row.fee} · {row.qt_label}"
        )
        self.fields["container"].setText(_text(row.container))
        self.fields["invoice"].setText(_text(row.invoice_no))
        self.fields["before"].setText(_money(row.value_before))
        self.fields["after"].setText(_money(row.value_after))
        location = (
            f"{row.sheet_name}, dòng {row.target_row}, ô {row.target_cell}"
            if row.target_row is not None and row.target_cell
            else "Chưa xác định vị trí BK"
        )
        self.fields["location"].setText(location)
        self.fields["source"].setText(_text(row.source_document_name))
        explanation = row.reason
        if row.source_count > 1:
            explanation += f" Gồm {row.source_count} khoản nguồn cùng ô BK."
        if row.source_row is not None:
            explanation += f" Dòng nguồn {row.source_row}."
        self.fields["reason"].setText(explanation.strip() or "—")


__all__ = ["PostingConfirmationDetails", "PostingDetailData", "PostingDetailRow"]
