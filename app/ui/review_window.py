"""Cửa sổ xem, kiểm tra và chỉnh sửa một batch JSON."""

from __future__ import annotations

import json
import logging
from collections.abc import Callable, Mapping, Sequence
from dataclasses import asdict, is_dataclass
from datetime import datetime
from typing import Any

from PySide6.QtCore import QModelIndex, Qt, Signal
from PySide6.QtGui import QCloseEvent, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .edit_row_dialog import EditRowDialog
from .inline_action_delegate import InlineActionDelegate
from .review_table_model import (
    ContainerLoadPresentation,
    FEE_CATALOG,
    ReviewFilterProxyModel,
    ReviewRow,
    ReviewStats,
    ReviewTableModel,
    RowStatus,
    coerce_review_row,
)
from app.constants import SCHEMA_VERSION
from app.models import DataRow
from app.services.review_carrier_lookup import ReviewCarrierLookup
from app.sea_freight.contracts import (
    GroupStatus,
    InvoiceHistoryMatchKind,
    group_status_text,
)
from app.sea_freight.service import VesselVoyageNotFoundError
from app.sea_freight.matching import normalize_match_key
from app.ui.sea_freight_center import ReconciliationPeriodDialog
from app.ui.feedback import LinearLoadingBar, set_button_loading

LOGGER = logging.getLogger(__name__)

STATUS_VI = {
    "RECEIVED": "Đã tiếp nhận",
    "REVIEWING": "Đang kiểm tra",
    "READY": "Đã xác nhận",
    "INVALID": "Không hợp lệ",
    "ARCHIVED": "Đã lưu trữ",
}


class SeaFreightContributionSelectionDialog(QDialog):
    """Chọn nhiều HĐ cùng tàu/chuyến để đưa vào một hồ sơ đối soát."""

    def __init__(
        self,
        candidates: Sequence[tuple[int, ReviewRow]],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Chọn hóa đơn cho hồ sơ cước biển")
        self.resize(850, 430)
        layout = QVBoxLayout(self)
        note = QLabel(
            "Các hóa đơn dưới đây cùng tàu/chuyến và chưa thuộc hồ sơ khác. "
            "Chọn các hóa đơn cần cộng số container trong lần này."
        )
        note.setWordWrap(True)
        layout.addWidget(note)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(
            ["Chứng từ", "Số HĐ", "Ngày HĐ", "SL cont", "Số tiền"]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self._items: list[QTableWidgetItem] = []
        for source_index, row_data in candidates:
            row = self.table.rowCount()
            self.table.insertRow(row)
            item = QTableWidgetItem(str(row_data.source_document_name))
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            item.setData(Qt.ItemDataRole.UserRole, source_index)
            self.table.setItem(row, 0, item)
            self.table.setItem(row, 1, QTableWidgetItem(str(row_data.invoice_no or "—")))
            self.table.setItem(row, 2, QTableWidgetItem(str(row_data.invoice_date or "—")))
            self.table.setItem(
                row, 3, QTableWidgetItem(str(row_data.invoice_container_count or "—"))
            )
            self.table.setItem(
                row,
                4,
                QTableWidgetItem(
                    f"{row_data.amount:,}".replace(",", ".")
                    if type(row_data.amount) is int
                    else "—"
                ),
            )
            self._items.append(item)
        layout.addWidget(self.table, 1)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Tạo / bổ sung hồ sơ")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    @property
    def selected_source_indices(self) -> list[int]:
        return [
            int(item.data(Qt.ItemDataRole.UserRole))
            for item in self._items
            if item.checkState() == Qt.CheckState.Checked
        ]


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
            if hasattr(value, "value"):
                value = value.value
            return value
    return default


def _format_datetime(value: Any) -> str:
    if value in (None, ""):
        return "—"
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y %H:%M:%S")
    text = str(value)
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        return parsed.strftime("%d/%m/%Y %H:%M:%S")
    except (ValueError, TypeError):
        return text


def _extract_review(batch: Any, explicit_rows: Any | None) -> tuple[Any, list[Any]]:
    metadata = _value(batch, "metadata", default=batch)
    if explicit_rows is not None:
        return metadata, list(explicit_rows)
    document = _value(batch, "document")
    if document is not None:
        rows = _value(document, "rows", "d", "data", default=[])
        return metadata, list(rows or [])
    rows = _value(batch, "rows", "d", "data")
    if rows is not None:
        return metadata, list(rows)
    if isinstance(batch, Sequence) and not isinstance(batch, (str, bytes, bytearray)):
        return None, list(batch)
    return metadata, []


class RawJsonDialog(QDialog):
    """Preview JSON chỉ đọc, không cho sửa lệch schema."""

    def __init__(self, document: dict[str, Any], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("JSON thô – chỉ đọc")
        self.resize(780, 610)
        layout = QVBoxLayout(self)
        note = QLabel(
            "Nội dung bên dưới là bản xem trước theo đúng schema v/d. "
            "Hãy dùng hộp sửa dòng để thay đổi dữ liệu."
        )
        note.setWordWrap(True)
        note.setProperty("muted", True)
        layout.addWidget(note)
        editor = QPlainTextEdit()
        editor.setObjectName("rawJsonView")
        editor.setReadOnly(True)
        editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        editor.setPlainText(json.dumps(document, ensure_ascii=False, indent=2))
        layout.addWidget(editor, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.button(QDialogButtonBox.StandardButton.Close).setText("Đóng")
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


class VesselVoyageNotFoundDialog(QDialog):
    EDIT_RESULT = 2

    def __init__(
        self,
        error: VesselVoyageNotFoundError,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Không tìm thấy tàu/chuyến")
        self.setModal(True)
        self.setMinimumWidth(560)
        layout = QVBoxLayout(self)
        title = QLabel("Không tìm thấy tàu/chuyến tương ứng trong BK")
        title.setObjectName("dialogTitle")
        layout.addWidget(title)
        message = QLabel(
            f'Không tìm thấy “{error.vessel_voyage}” trong sheet {error.bk_sheet}.\n'
            "Hãy kiểm tra, sửa lại tên tàu hoặc số chuyến rồi bấm Đối soát lại."
        )
        message.setWordWrap(True)
        layout.addWidget(message)
        if error.suggestions:
            suggestion_title = QLabel("Tàu/chuyến có thể tương ứng trong BK:")
            suggestion_title.setStyleSheet("font-weight: 600;")
            layout.addWidget(suggestion_title)
            suggestion_text = QLabel(
                "\n".join(
                    f"• {item.vessel_voyage} — {item.container_count} container"
                    for item in error.suggestions
                )
            )
            suggestion_text.setObjectName("vesselVoyageSuggestions")
            suggestion_text.setTextInteractionFlags(
                Qt.TextInteractionFlag.TextSelectableByMouse
            )
            layout.addWidget(suggestion_text)
        buttons = QDialogButtonBox()
        self.close_button = buttons.addButton(
            "Đóng", QDialogButtonBox.ButtonRole.RejectRole
        )
        self.edit_button = buttons.addButton(
            "Sửa dòng", QDialogButtonBox.ButtonRole.AcceptRole
        )
        self.edit_button.setProperty("primary", True)
        self.close_button.clicked.connect(self.reject)
        self.edit_button.clicked.connect(lambda: self.done(self.EDIT_RESULT))
        layout.addWidget(buttons)


class ReviewWindow(QMainWindow):
    """Cửa sổ review hoàn chỉnh, nhận service/callback qua dependency injection."""

    saveRequested = Signal(object, object)
    confirmRequested = Signal(object, object)
    saved = Signal(object)
    confirmed = Signal(object)
    batchUpdated = Signal(object)
    reconciliationChanged = Signal()
    reconciliationOpenRequested = Signal(int)
    closed = Signal()

    def __init__(
        self,
        batch: Any = None,
        rows: Any | None = None,
        parent: QWidget | None = None,
        *,
        batch_service: Any | None = None,
        validator: Any | None = None,
        save_handler: Callable[[Any, Any], Any] | None = None,
        confirm_handler: Callable[[Any, Any], Any] | None = None,
        sea_freight_service: Any | None = None,
        settings: Any | None = None,
    ) -> None:
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self.setWindowTitle("Xem và chỉnh sửa dữ liệu bóc tách")
        self.setMinimumSize(920, 600)
        self.resize(1120, 720)
        self._batch_service = batch_service
        self._validator = validator
        self._save_handler = save_handler
        self._confirm_handler = confirm_handler
        self._sea_freight_service = sea_freight_service
        self._settings = settings
        self._metadata, initial_rows = _extract_review(batch, rows)
        self._batch_id = _value(self._metadata, "id", "batch_id")
        self._last_saved_at = _value(self._metadata, "last_saved_at")
        self._status = str(_value(self._metadata, "status", default="REVIEWING"))
        self._status = self._status.split(".")[-1].upper()
        self._saving = False
        self._allow_negative = (
            str(_value(self._metadata, "source_kind", default="ASSISTANT"))
            == "BANG_KE"
        )
        initial_rows, populated_carrier_count = self._rows_with_standard_carriers(
            initial_rows
        )

        self.model = ReviewTableModel(
            initial_rows,
            validator=validator,
            allow_negative=self._allow_negative,
        )
        self._source_index_by_runtime = {
            self.model.runtime_id_at(index): index
            for index in range(self.model.rowCount())
        }
        self.proxy_model = ReviewFilterProxyModel(self)
        self.proxy_model.setSourceModel(self.model)
        self._build_ui()
        self._refresh_document_filter()
        self._connect_signals()
        self._install_shortcuts()
        self._update_metadata_labels()
        self._update_stats(self.model.stats)
        self._update_dirty(False)
        self._update_action_state()
        self._restore_reconciliation_presentations()
        if populated_carrier_count:
            self.model.mark_dirty()
            self.statusBar().showMessage(
                f"Đã điền Bên vận tải chuẩn từ file Hàng ngày cho "
                f"{populated_carrier_count} khoản chi; hãy kiểm tra và xác nhận.",
                12000,
            )

    def _rows_with_standard_carriers(
        self, rows: Sequence[Any]
    ) -> tuple[list[ReviewRow], int]:
        prepared = [coerce_review_row(row) for row in rows]
        source_kind = str(
            _value(self._metadata, "source_kind", default="ASSISTANT")
        ).split(".")[-1].upper()
        if source_kind != "ASSISTANT":
            return prepared, 0
        daily_path = _value(self._settings, "daily_workbook_path", default="")
        if not str(daily_path or "").strip():
            return prepared, 0
        try:
            resolved = ReviewCarrierLookup(daily_path).resolve(prepared)
        except Exception:
            LOGGER.exception("Không thể tra Bên vận tải chuẩn cho màn hình review.")
            return prepared, 0
        for index, carrier in resolved.items():
            prepared[index].carrier = carrier
        return prepared, len(resolved)

    def _build_ui(self) -> None:
        central = QWidget()
        central.setObjectName("applicationRoot")
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)

        top_line = QHBoxLayout()
        title_box = QVBoxLayout()
        title = QLabel("Kiểm tra dữ liệu bóc tách")
        title.setObjectName("pageTitle")
        title_box.addWidget(title)
        self.source_label = QLabel()
        self.source_label.setProperty("muted", True)
        title_box.addWidget(self.source_label)
        top_line.addLayout(title_box, 1)
        self.dirty_label = QLabel("Có thay đổi chưa lưu")
        self.dirty_label.setObjectName("dirtyLabel")
        self.dirty_label.setStyleSheet(
            "color: #A16207; background: #FFF8DB; border-radius: 6px; "
            "padding: 6px 10px; font-weight: 600;"
        )
        top_line.addWidget(self.dirty_label, 0, Qt.AlignmentFlag.AlignTop)
        root.addLayout(top_line)

        meta_card = QFrame()
        meta_card.setProperty("card", True)
        meta_layout = QGridLayout(meta_card)
        meta_layout.setContentsMargins(12, 7, 12, 7)
        meta_layout.setHorizontalSpacing(16)
        meta_layout.setVerticalSpacing(2)
        self.batch_id_value = QLabel()
        self.sha_value = QLabel()
        self.received_value = QLabel()
        self.saved_value = QLabel()
        self.status_value = QLabel()
        metadata = (
            ("Batch ID", self.batch_id_value),
            ("SHA-256", self.sha_value),
            ("Thời điểm nhận", self.received_value),
            ("Lưu gần nhất", self.saved_value),
            ("Trạng thái", self.status_value),
        )
        for column, (label, value) in enumerate(metadata):
            caption = QLabel(label)
            caption.setProperty("muted", True)
            value.setStyleSheet("font-weight: 600;")
            meta_layout.addWidget(caption, 0, column)
            meta_layout.addWidget(value, 1, column)
        root.addWidget(meta_card)

        stats_line = QHBoxLayout()
        stats_line.setSpacing(8)
        self.stat_labels: dict[str, QLabel] = {}
        stat_defs = (
            ("total", "Tổng dòng"),
            ("valid", "Hợp lệ"),
            ("warning", "Cảnh báo"),
            ("error", "Lỗi"),
            ("amount", "Tổng tiền"),
        )
        for key, caption in stat_defs:
            card = QFrame()
            card.setProperty("card", True)
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(10, 5, 10, 5)
            card_layout.setSpacing(2)
            cap_label = QLabel(caption)
            cap_label.setProperty("muted", True)
            value_label = QLabel("0")
            value_label.setObjectName(f"{key}Stat")
            value_label.setStyleSheet("font-size: 12pt; font-weight: 700;")
            card_layout.addWidget(cap_label)
            card_layout.addWidget(value_label)
            stats_line.addWidget(card, 1)
            self.stat_labels[key] = value_label
        root.addLayout(stats_line)

        filter_toolbar = QHBoxLayout()
        filter_toolbar.setSpacing(7)
        self.search_edit = QLineEdit()
        self.search_edit.setObjectName("reviewSearchEdit")
        self.search_edit.setPlaceholderText(
            "Tìm container, B/L, Số HĐ hoặc Bên vận tải…  (Ctrl+F)"
        )
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setMinimumWidth(220)
        filter_toolbar.addWidget(self.search_edit, 2)

        self.document_filter = QComboBox()
        self.document_filter.setObjectName("documentFilter")
        self.document_filter.setMinimumWidth(210)
        filter_toolbar.addWidget(self.document_filter, 2)

        self.fee_filter = QComboBox()
        self.fee_filter.setObjectName("feeFilter")
        self.fee_filter.addItem("Tất cả loại cước", "")
        for code, name in FEE_CATALOG.items():
            self.fee_filter.addItem(f"{code} – {name}", code)
        filter_toolbar.addWidget(self.fee_filter, 2)

        self.status_filter = QComboBox()
        self.status_filter.setObjectName("statusFilter")
        self.status_filter.addItem("Tất cả trạng thái", "")
        self.status_filter.addItem("Hợp lệ", RowStatus.VALID.value)
        self.status_filter.addItem("Cảnh báo", RowStatus.WARNING.value)
        self.status_filter.addItem("Lỗi", RowStatus.ERROR.value)
        filter_toolbar.addWidget(self.status_filter, 1)

        self.clear_filter_button = QPushButton("Xóa bộ lọc")
        self.clear_filter_button.setObjectName("clearFilterButton")
        filter_toolbar.addWidget(self.clear_filter_button)
        root.addLayout(filter_toolbar)

        action_toolbar = QHBoxLayout()
        action_toolbar.setSpacing(7)
        action_toolbar.addStretch(1)
        self.add_button = QPushButton("Thêm dòng")
        self.add_button.setObjectName("addRowButton")
        self.edit_button = QPushButton("Sửa dòng")
        self.edit_button.setObjectName("editRowButton")
        self.delete_button = QPushButton("Xóa dòng")
        self.delete_button.setObjectName("deleteRowButton")
        self.delete_button.setProperty("danger", True)
        self.raw_button = QPushButton("Xem JSON thô")
        self.raw_button.setObjectName("rawJsonButton")
        for button in (self.add_button, self.edit_button, self.delete_button, self.raw_button):
            action_toolbar.addWidget(button)
        root.addLayout(action_toolbar)

        self.table = QTableView()
        self.table.setObjectName("reviewTable")
        self.table.setModel(self.proxy_model)
        for column in (
            ReviewTableModel.COLUMN_CONTAINER_COUNT_BASIS,
            ReviewTableModel.COLUMN_FEE,
            ReviewTableModel.COLUMN_RULE,
            ReviewTableModel.COLUMN_RULE_NAME,
            ReviewTableModel.COLUMN_STATUS,
        ):
            self.table.setColumnHidden(column, True)
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setToolTip(
            "Giữ Ctrl để chọn từng dòng hoặc Shift để chọn một dải dòng."
        )
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setWordWrap(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(32)
        header = self.table.horizontalHeader()
        header.setSectionsClickable(True)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(64)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(ReviewTableModel.COLUMN_NO, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(ReviewTableModel.COLUMN_FEE_NAME, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(ReviewTableModel.COLUMN_MESSAGES, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(ReviewTableModel.COLUMN_CONT, 140)
        self.table.setColumnWidth(ReviewTableModel.COLUMN_BL, 130)
        self.table.setColumnWidth(ReviewTableModel.COLUMN_VESSEL_VOYAGE, 190)
        self.table.setColumnWidth(ReviewTableModel.COLUMN_INVOICE_NO, 130)
        self.table.setColumnWidth(ReviewTableModel.COLUMN_CARRIER, 190)
        self.table.setColumnWidth(ReviewTableModel.COLUMN_AMOUNT, 185)
        self.table.setColumnWidth(ReviewTableModel.COLUMN_LOOKUP_ACTION, 105)
        self.lookup_action_delegate = InlineActionDelegate(self.table)
        self.table.setItemDelegateForColumn(
            ReviewTableModel.COLUMN_LOOKUP_ACTION,
            self.lookup_action_delegate,
        )
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        root.addWidget(self.table, 1)

        self.save_loading_bar = LinearLoadingBar()
        self.save_loading_bar.setAccessibleName("Tiến trình lưu dữ liệu")
        root.addWidget(self.save_loading_bar)

        bottom = QHBoxLayout()
        bottom.setSpacing(8)
        self.visible_label = QLabel()
        self.visible_label.setProperty("muted", True)
        bottom.addWidget(self.visible_label)
        bottom.addStretch(1)
        self.confirm_button = QPushButton("Lưu")
        self.confirm_button.setObjectName("confirmButton")
        self.confirm_button.setProperty("primary", True)
        self.close_button = QPushButton("Đóng")
        self.close_button.setObjectName("closeReviewButton")
        self.confirm_button.setMinimumWidth(110)
        self.close_button.setMinimumWidth(110)
        bottom.addWidget(self.close_button)
        bottom.addWidget(self.confirm_button)
        root.addLayout(bottom)

    def _connect_signals(self) -> None:
        self.search_edit.textChanged.connect(self.proxy_model.set_search_text)
        self.search_edit.textChanged.connect(self._update_visible_count)
        self.document_filter.currentIndexChanged.connect(
            lambda: self.proxy_model.set_document_filter(
                self.document_filter.currentData()
            )
        )
        self.document_filter.currentIndexChanged.connect(self._update_visible_count)
        self.fee_filter.currentIndexChanged.connect(
            lambda: self.proxy_model.set_fee_filter(self.fee_filter.currentData())
        )
        self.fee_filter.currentIndexChanged.connect(self._update_visible_count)
        self.status_filter.currentIndexChanged.connect(
            lambda: self.proxy_model.set_status_filter(self.status_filter.currentData())
        )
        self.status_filter.currentIndexChanged.connect(self._update_visible_count)
        self.clear_filter_button.clicked.connect(self.clear_filters)
        self.add_button.clicked.connect(self.add_row)
        self.edit_button.clicked.connect(self.edit_selected_row)
        self.delete_button.clicked.connect(self.delete_selected_row)
        self.raw_button.clicked.connect(self.show_raw_json)
        self.confirm_button.clicked.connect(self.confirm_batch)
        self.close_button.clicked.connect(self.close)
        self.table.doubleClicked.connect(self.edit_selected_row)
        self.table.selectionModel().selectionChanged.connect(self._update_action_state)
        self.model.dirtyChanged.connect(self._dirty_state_changed)
        self.model.validationChanged.connect(self._update_stats)
        self.model.rowsChanged.connect(self._update_visible_count)
        self.model.rowsChanged.connect(self._refresh_document_filter)
        self.model.rowsChanged.connect(self._restore_reconciliation_presentations)
        self.lookup_action_delegate.clicked.connect(
            self._lookup_action_clicked
        )

    def _install_shortcuts(self) -> None:
        self.save_shortcut = QShortcut(QKeySequence.StandardKey.Save, self)
        self.save_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self.save_shortcut.activated.connect(self.confirm_batch)

        self.find_shortcut = QShortcut(QKeySequence.StandardKey.Find, self)
        self.find_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self.find_shortcut.activated.connect(self._focus_search)

        self.delete_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), self.table)
        self.delete_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.delete_shortcut.activated.connect(self.delete_selected_row)

        self.enter_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Return), self.table)
        self.enter_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self.enter_shortcut.activated.connect(self.edit_selected_row)

    def _focus_search(self) -> None:
        self.search_edit.setFocus(Qt.FocusReason.ShortcutFocusReason)
        self.search_edit.selectAll()

    def _update_metadata_labels(self) -> None:
        filename = _value(
            self._metadata, "source_filename", "filename", "file_name", default="Chưa có tên file"
        )
        self.source_label.setText(str(filename))
        self.batch_id_value.setText("—" if self._batch_id is None else str(self._batch_id))
        sha = str(_value(self._metadata, "sha256", "hash", default="") or "")
        self.sha_value.setText(f"{sha[:12]}…" if len(sha) > 12 else (sha or "—"))
        self.sha_value.setToolTip(sha)
        self.received_value.setText(
            _format_datetime(_value(self._metadata, "received_at", "created_at"))
        )
        self.saved_value.setText(_format_datetime(self._last_saved_at))
        self.status_value.setText(STATUS_VI.get(self._status, self._status or "—"))

    def _update_stats(self, stats: ReviewStats) -> None:
        self._global_stats = stats
        self._update_filtered_stats()
        self.confirm_button.setEnabled(stats.error == 0 and not self._saving)
        self._update_visible_count()

    def _set_stat_labels(self, stats: ReviewStats) -> None:
        self.stat_labels["total"].setText(f"{stats.total:,}".replace(",", "."))
        self.stat_labels["valid"].setText(f"{stats.valid:,}".replace(",", "."))
        self.stat_labels["warning"].setText(f"{stats.warning:,}".replace(",", "."))
        self.stat_labels["error"].setText(f"{stats.error:,}".replace(",", "."))
        self.stat_labels["amount"].setText(f"{stats.total_amount:,} ₫".replace(",", "."))
        self.stat_labels["valid"].setStyleSheet("font-size: 12pt; font-weight: 700; color: #15803D;")
        self.stat_labels["warning"].setStyleSheet("font-size: 12pt; font-weight: 700; color: #A16207;")
        self.stat_labels["error"].setStyleSheet("font-size: 12pt; font-weight: 700; color: #B42318;")

    def _update_filtered_stats(self) -> None:
        source_rows = [
            self.proxy_model.mapToSource(self.proxy_model.index(index, 0)).row()
            for index in range(self.proxy_model.rowCount())
        ]
        validations = [self.model.validation_at(index) for index in source_rows]
        rows = [self.model.row_at(index) for index in source_rows]
        stats = ReviewStats(
            total=len(rows),
            valid=sum(item.status is RowStatus.VALID for item in validations),
            warning=sum(item.status is RowStatus.WARNING for item in validations),
            error=sum(item.status is RowStatus.ERROR for item in validations),
            with_container=sum(bool(item.cont) for item in rows),
            with_bl=sum(bool(item.bl) for item in rows),
            with_amount=sum(type(item.amount) is int for item in rows),
            total_amount=sum(item.amount for item in rows if type(item.amount) is int),
            fee_counts={},
        )
        self._set_stat_labels(stats)

    def _update_visible_count(self, *_args: Any) -> None:
        self._update_filtered_stats()
        self.visible_label.setText(
            f"Đang hiển thị {self.proxy_model.rowCount():,}/{self.model.rowCount():,} dòng".replace(
                ",", "."
            )
        )

    def _update_dirty(self, dirty: bool) -> None:
        self.dirty_label.setVisible(dirty)
        suffix = " *" if dirty else ""
        self.setWindowTitle(f"Xem và chỉnh sửa dữ liệu bóc tách{suffix}")

    def _dirty_state_changed(self, dirty: bool) -> None:
        self._update_dirty(dirty)
        if dirty and self._status == "READY":
            self._reopen_ready_batch()

    def _reopen_ready_batch(self) -> None:
        """Chuyển READY về REVIEWING ngay khi người dùng thực sự sửa dữ liệu."""

        handler = self._service_handler("reopen_batch", "reopen")
        if handler is None or self._batch_id is None:
            return
        try:
            result = handler(self._batch_id)
            metadata = _value(result, "metadata", default=result)
            if metadata is not None:
                self._metadata = metadata
                self._status = str(
                    _value(metadata, "status", default="REVIEWING")
                ).split(".")[-1].upper()
                self._last_saved_at = _value(
                    metadata, "last_saved_at", default=self._last_saved_at
                )
                self._update_metadata_labels()
                self.batchUpdated.emit(result)
        except Exception as exc:
            # Không làm mất thao tác vừa sửa. Lần lưu tiếp theo vẫn bắt buộc gọi
            # service và sẽ báo lỗi đầy đủ nếu database thực sự không cập nhật được.
            self.statusBar().showMessage(
                f"Chưa cập nhật được trạng thái batch: {exc}", 8000
            )

    def _update_action_state(self, *_args: Any) -> None:
        selected_count = (
            len(self.table.selectionModel().selectedRows())
            if self.table.model()
            else 0
        )
        self.edit_button.setEnabled(selected_count == 1)
        self.delete_button.setEnabled(selected_count > 0)
        self.delete_button.setText(
            f"Xóa {selected_count} dòng" if selected_count > 1 else "Xóa dòng"
        )

    def clear_filters(self) -> None:
        self.search_edit.clear()
        self.document_filter.setCurrentIndex(0)
        self.fee_filter.setCurrentIndex(0)
        self.status_filter.setCurrentIndex(0)
        self.proxy_model.clear_filters()
        self._update_visible_count()

    def _refresh_document_filter(self, *_args: Any) -> None:
        selected = self.document_filter.currentData() if self.document_filter.count() else ""
        documents: dict[str, tuple[str, int]] = {}
        for row in self.model.rows():
            name, count = documents.get(
                row.source_document_id, (row.source_document_name, 0)
            )
            documents[row.source_document_id] = (name, count + 1)
        self.document_filter.blockSignals(True)
        self.document_filter.clear()
        self.document_filter.addItem(
            f"Tất cả chứng từ – {self.model.rowCount()} dòng", ""
        )
        for document_id, (name, count) in documents.items():
            self.document_filter.addItem(f"{name} – {count} dòng", document_id)
        index = self.document_filter.findData(selected)
        self.document_filter.setCurrentIndex(max(0, index))
        self.document_filter.blockSignals(False)
        self.proxy_model.set_document_filter(self.document_filter.currentData())

    def _selected_source_row(self) -> int | None:
        selected = self._selected_source_rows()
        return selected[0] if len(selected) == 1 else None

    def _selected_source_rows(self) -> list[int]:
        selected = self.table.selectionModel().selectedRows()
        source_rows = {
            source_index.row()
            for proxy_index in selected
            if (source_index := self.proxy_model.mapToSource(proxy_index)).isValid()
        }
        return sorted(source_rows)

    def _select_source_row(self, source_row: int) -> None:
        source_index = self.model.index(source_row, ReviewTableModel.COLUMN_NO)
        proxy_index = self.proxy_model.mapFromSource(source_index)
        if not proxy_index.isValid():
            self.clear_filters()
            proxy_index = self.proxy_model.mapFromSource(source_index)
        if proxy_index.isValid():
            self.table.selectRow(proxy_index.row())
            self.table.scrollTo(proxy_index, QAbstractItemView.ScrollHint.PositionAtCenter)

    def add_row(self) -> None:
        source_document_id = str(self.document_filter.currentData() or "")
        source_document_name = ""
        if source_document_id:
            source_document_name = next(
                (
                    row.source_document_name
                    for row in self.model.rows()
                    if row.source_document_id == source_document_id
                ),
                "",
            )
        else:
            choices = [
                f"{self.document_filter.itemText(index)} [{self.document_filter.itemData(index)}]"
                for index in range(1, self.document_filter.count())
            ]
            choices.append("Dòng thêm thủ công [MANUAL]")
            choice, accepted = QInputDialog.getItem(
                self,
                "Chọn chứng từ nguồn",
                "Dòng mới thuộc chứng từ nào?",
                choices,
                len(choices) - 1,
                False,
            )
            if not accepted:
                return
            if choice == "Dòng thêm thủ công [MANUAL]":
                source_document_id = "MANUAL"
                source_document_name = "Dòng thêm thủ công"
            else:
                selected_index = choices.index(choice) + 1
                source_document_id = str(self.document_filter.itemData(selected_index))
                source_document_name = next(
                    row.source_document_name
                    for row in self.model.rows()
                    if row.source_document_id == source_document_id
                )
        dialog = EditRowDialog(
            parent=self,
            validator=self._validator,
            allow_negative=self._allow_negative,
            source_document_id=source_document_id,
            source_document_name=source_document_name,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        position = self.model.add_row(dialog.row_data())
        self._select_source_row(position)

    def edit_selected_row(self, _index: QModelIndex | None = None) -> None:
        if _index is not None and _index.isValid() and _index.column() in {
            ReviewTableModel.COLUMN_LOOKUP_ACTION,
        }:
            return
        source_row = self._selected_source_row()
        if source_row is None:
            return
        self._edit_source_row(source_row)

    def _edit_source_row(self, source_row: int, *, focus_vessel: bool = False) -> bool:
        runtime_id = self.model.runtime_id_at(source_row)
        if self.model.lookup_presentation(runtime_id).session_id:
            QMessageBox.information(
                self,
                "Dòng đã có hồ sơ đối soát",
                "Hãy bấm Mở hồ sơ và sửa HĐ trực tiếp trong cửa sổ Đối soát số cont.",
            )
            return False
        dialog = EditRowDialog(
            self.model.row_at(source_row),
            parent=self,
            validator=self._validator,
            allow_negative=self._allow_negative,
        )
        if focus_vessel:
            dialog.focus_vessel_fields()
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return False
        self.model.update_row(source_row, dialog.row_data())
        self._select_source_row(source_row)
        return True

    def delete_selected_row(self) -> None:
        source_rows = self._selected_source_rows()
        if not source_rows:
            return
        if any(
            self.model.lookup_presentation(self.model.runtime_id_at(index)).session_id
            for index in range(self.model.rowCount())
        ):
            QMessageBox.information(
                self,
                "Có HĐ đang đối soát",
                "Không thể xóa dòng vì sẽ làm lệch vị trí nguồn của nhóm chờ. "
                "Hãy mở hồ sơ đối soát và xóa HĐ tại đó.",
            )
            return
        if len(source_rows) == 1:
            confirmation = (
                f"Bạn có chắc muốn xóa dòng số {source_rows[0] + 1}? Dòng sẽ chỉ bị xóa "
                "khỏi bản làm việc sau khi bạn lưu."
            )
        else:
            display_numbers = [str(row + 1) for row in source_rows[:10]]
            if len(source_rows) > 10:
                display_numbers.append("…")
            confirmation = (
                f"Bạn có chắc muốn xóa {len(source_rows)} dòng đã chọn "
                f"(STT {', '.join(display_numbers)})? Các dòng sẽ chỉ bị xóa "
                "khỏi bản làm việc sau khi bạn lưu."
            )
        selected_proxy_rows = [
            index.row() for index in self.table.selectionModel().selectedRows()
        ]
        next_proxy_row = min(selected_proxy_rows)
        answer = QMessageBox.question(
            self,
            "Xóa dữ liệu đã chọn",
            confirmation,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.model.remove_rows(source_rows)
        if self.proxy_model.rowCount():
            self.table.selectRow(
                min(next_proxy_row, self.proxy_model.rowCount() - 1)
            )

    def show_raw_json(self) -> None:
        RawJsonDialog(self.model.to_document(), self).exec()

    def set_save_handler(self, handler: Callable[[Any, Any], Any] | None) -> None:
        self._save_handler = handler

    def set_confirm_handler(self, handler: Callable[[Any, Any], Any] | None) -> None:
        self._confirm_handler = handler

    def _service_handler(self, *names: str) -> Callable[..., Any] | None:
        for owner in (self._batch_service,):
            if owner is None:
                continue
            for name in names:
                method = getattr(owner, name, None)
                if callable(method):
                    return method
        return None

    def _core_document(self) -> Any:
        objects = [row.to_object() for row in self.model.rows()]
        try:
            from app.models import BatchDocument, DataRow

            rows = [DataRow.from_mapping(row) for row in objects]
            try:
                return BatchDocument(v=SCHEMA_VERSION, rows=rows)
            except TypeError:
                return BatchDocument(version=SCHEMA_VERSION, data=rows)
        except (ImportError, AttributeError, TypeError):
            return {"v": SCHEMA_VERSION, "d": objects}

    def save_working(self) -> bool:
        if self._saving:
            return False
        handler = self._save_handler or self._service_handler(
            "save_working", "save_batch", "save_working_copy"
        )
        if handler is None:
            QMessageBox.warning(
                self,
                "Chưa thể lưu",
                "Chưa kết nối dịch vụ lưu bản làm việc. Hãy đóng cửa sổ và thử lại "
                "sau khi ứng dụng khởi tạo xong.",
            )
            return False
        document = self._core_document()
        self.saveRequested.emit(self._batch_id, document)
        self._set_saving(True)
        try:
            result = handler(self._batch_id, document)
            if result is False:
                raise RuntimeError("Dịch vụ từ chối lưu bản làm việc.")
            self._apply_service_result(result)
            self.model.mark_clean()
            self._last_saved_at = (
                _value(self._metadata, "last_saved_at") or datetime.now()
            )
            self.saved_value.setText(_format_datetime(self._last_saved_at))
            self.saved.emit(result if result is not None else self._metadata)
            self.batchUpdated.emit(result if result is not None else self._metadata)
            self.statusBar().showMessage("Đã lưu an toàn bản đang chỉnh sửa.", 5000)
            return True
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Không lưu được dữ liệu",
                f"Ứng dụng chưa thể lưu bản làm việc: {exc}\n"
                "Dữ liệu đang sửa vẫn còn trong cửa sổ này; hãy thử lại.",
            )
            return False
        finally:
            self._set_saving(False)

    def confirm_batch(self) -> bool:
        stats = self.model.stats
        unmanaged = [
            index
            for index in range(self.model.rowCount())
            if self.model.row_at(index).fee == "CB"
            and self.model.row_at(index).cont in (None, "")
            and not self.model.lookup_presentation(
                self.model.runtime_id_at(index)
            ).session_id
        ]
        if unmanaged:
            self._select_source_row(unmanaged[0])
            QMessageBox.warning(
                self,
                "Cước biển chưa được đối soát",
                f"Còn {len(unmanaged)} dòng cước biển thiếu số cont chưa có hồ sơ đối soát. "
                "Hãy bổ sung tàu/chuyến, số cont rồi bấm Đối soát số cont.",
            )
            return False
        incomplete = [
            index
            for index in range(self.model.rowCount())
            if self.model.row_at(index).fee == "CB"
            and self.model.row_at(index).cont in (None, "")
            and self.model.lookup_presentation(
                self.model.runtime_id_at(index)
            ).session_id
            and self.model.lookup_presentation(
                self.model.runtime_id_at(index)
            ).status not in {"ALLOCATED", "POSTED"}
        ]
        if incomplete:
            self._select_source_row(incomplete[0])
            QMessageBox.warning(
                self,
                "Hồ sơ đối soát chưa hoàn tất",
                f"Còn {len(incomplete)} dòng cước biển có hồ sơ nhưng chưa xác nhận kết quả. "
                "Hãy bấm Xem hồ sơ và hoàn tất đối soát trước khi lưu file.",
            )
            return False
        if stats.error:
            self._focus_first_error()
            QMessageBox.warning(
                self,
                "Còn lỗi chặn",
                f"Batch còn {stats.error} dòng lỗi. Hãy sửa lỗi trước khi xác nhận hoàn tất.",
            )
            return False
        if stats.warning:
            answer = QMessageBox.question(
                self,
                "Lưu khi còn cảnh báo",
                f"Dữ liệu còn {stats.warning} dòng cảnh báo. Bạn đã kiểm tra và vẫn "
                "muốn lưu?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return False

        handler = self._confirm_handler or self._service_handler(
            "confirm_batch", "confirm", "mark_ready"
        )
        if handler is None:
            QMessageBox.warning(
                self,
                "Chưa thể lưu",
                "Chưa kết nối dịch vụ lưu dữ liệu.",
            )
            return False

        document = self._core_document()
        self.confirmRequested.emit(self._batch_id, document)
        self._set_saving(True)
        try:
            result = handler(self._batch_id, document)
            if result is False:
                raise RuntimeError("Dịch vụ từ chối xác nhận batch.")
            self._apply_service_result(result)
            self._status = "READY"
            self.status_value.setText(STATUS_VI["READY"])
            self.model.mark_clean()
            self.confirmed.emit(result if result is not None else self._metadata)
            self.batchUpdated.emit(result if result is not None else self._metadata)
            QMessageBox.information(
                self,
                "Lưu thành công",
                "Đã lưu thành công.",
            )
            return True
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Không lưu được dữ liệu",
                f"Không thể lưu dữ liệu: {exc}\nHãy kiểm tra log và thử lại.",
            )
            return False
        finally:
            self._set_saving(False)

    def _focus_first_error(self) -> None:
        source_row = self.model.first_error_row()
        if source_row is None:
            return
        self.clear_filters()
        self._select_source_row(source_row)

    def _set_saving(self, saving: bool) -> None:
        self._saving = saving
        self.save_loading_bar.set_running(saving)
        set_button_loading(self.confirm_button, saving)
        self.confirm_button.setText("Đang lưu…" if saving else "Lưu")
        self.confirm_button.setEnabled(not saving and self.model.stats.error == 0)
        self.add_button.setEnabled(not saving)
        self.edit_button.setEnabled(not saving and self._selected_source_row() is not None)
        self.delete_button.setEnabled(not saving and self._selected_source_row() is not None)

    def _apply_service_result(self, result: Any) -> None:
        if result is None:
            return
        metadata = _value(result, "metadata")
        document = _value(result, "document")
        if metadata is not None:
            self._metadata = metadata
            self._allow_negative = (
                str(_value(metadata, "source_kind", default="ASSISTANT"))
                == "BANG_KE"
            )
            self.model.set_allow_negative(self._allow_negative)
            self._batch_id = _value(metadata, "id", "batch_id", default=self._batch_id)
            status = _value(metadata, "status")
            if status is not None:
                self._status = str(status).split(".")[-1].upper()
            self._last_saved_at = _value(
                metadata, "last_saved_at", default=self._last_saved_at
            )
            self._update_metadata_labels()
        if document is not None:
            rows = _value(document, "rows", "d", "data")
            if rows is not None:
                self.model.set_rows(rows, mark_dirty=False)
                self._source_index_by_runtime = {
                    self.model.runtime_id_at(index): index
                    for index in range(self.model.rowCount())
                }
                self._restore_reconciliation_presentations()

    def replace_review(self, review: Any) -> None:
        """Nạp lại batch từ service mà không tái tạo cửa sổ."""

        if self.model.dirty:
            raise RuntimeError("Không thể nạp lại khi còn thay đổi chưa lưu.")
        metadata, rows = _extract_review(review, None)
        self._metadata = metadata
        self._allow_negative = (
            str(_value(metadata, "source_kind", default="ASSISTANT"))
            == "BANG_KE"
        )
        self.model.set_allow_negative(self._allow_negative)
        self._batch_id = _value(metadata, "id", "batch_id")
        self._status = str(_value(metadata, "status", default="REVIEWING")).split(".")[-1].upper()
        self._last_saved_at = _value(metadata, "last_saved_at")
        rows, populated_carrier_count = self._rows_with_standard_carriers(rows)
        self.model.set_rows(rows, mark_dirty=bool(populated_carrier_count))
        if populated_carrier_count:
            self.statusBar().showMessage(
                f"Đã điền Bên vận tải chuẩn từ file Hàng ngày cho "
                f"{populated_carrier_count} khoản chi; hãy kiểm tra và xác nhận.",
                12000,
            )
        self._update_metadata_labels()
        self._source_index_by_runtime = {
            self.model.runtime_id_at(index): index
            for index in range(self.model.rowCount())
        }

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802
        if self.model.dirty:
            message = QMessageBox(self)
            message.setIcon(QMessageBox.Icon.Warning)
            message.setWindowTitle("Có thay đổi chưa lưu")
            message.setText("Bản đang chỉnh sửa có thay đổi chưa được lưu.")
            message.setInformativeText("Bạn muốn lưu trước khi đóng không?")
            save = message.addButton("Lưu", QMessageBox.ButtonRole.AcceptRole)
            discard = message.addButton("Không lưu", QMessageBox.ButtonRole.DestructiveRole)
            cancel = message.addButton("Hủy đóng", QMessageBox.ButtonRole.RejectRole)
            message.setDefaultButton(save)
            message.exec()
            clicked = message.clickedButton()
            if clicked is cancel:
                event.ignore()
                return
            if clicked is save and not self.confirm_batch():
                event.ignore()
                return
            if clicked is not discard and clicked is not save:
                event.ignore()
                return
        self.closed.emit()
        event.accept()

    def _restore_reconciliation_presentations(self) -> None:
        if self._sea_freight_service is None or self._batch_id is None:
            return
        service = self._sea_freight_service
        batch_id = int(self._batch_id)
        source_sha256 = str(_value(self._metadata, "sha256", default="") or "")
        indexed_rows: list[tuple[int, DataRow]] = []
        runtime_by_source_index: dict[int, str] = {}
        for model_index in range(self.model.rowCount()):
            runtime_id = self.model.runtime_id_at(model_index)
            source_index = self._source_index_by_runtime.get(runtime_id, model_index)
            runtime_by_source_index[source_index] = runtime_id
            indexed_rows.append(
                (
                    source_index,
                    DataRow.from_mapping(self.model.row_at(model_index).to_object()),
                )
            )
        matches: dict[int, Any] = {}
        try:
            matches = service.sync_batch_history(
                indexed_rows,
                source_batch_id=batch_id,
                source_sha256=source_sha256,
            )
            managed = service.repository.groups_for_source_batch(batch_id)
        except Exception:
            LOGGER.exception("Không thể khôi phục trạng thái đối soát của batch %s", self._batch_id)
            try:
                managed = service.repository.groups_for_source_batch(batch_id)
            except Exception:
                managed = {}

        warnings: dict[str, tuple[str, ...]] = {}
        for source_index, match in matches.items():
            runtime_id = runtime_by_source_index.get(source_index)
            if runtime_id is None or match.group is None:
                continue
            if match.kind is InvoiceHistoryMatchKind.EXACT:
                invoice_no = getattr(match.contribution, "invoice_no", None) or "—"
                warnings[runtime_id] = (
                    f"HĐ {invoice_no} đã có trong hồ sơ #{match.group.id} – "
                    f"Lần {match.group.revision_no} – {group_status_text(match.group)}; "
                    "vẫn được phép ghi BK lại.",
                )
            elif match.kind is InvoiceHistoryMatchKind.CONFLICT:
                invoice_no = getattr(match.contribution, "invoice_no", None) or "—"
                differences = ", ".join(match.differing_fields)
                warnings[runtime_id] = (
                    f"HĐ {invoice_no} đã xuất hiện trong hồ sơ #{match.group.id} nhưng "
                    f"khác {differences}; hãy kiểm tra trước khi tiếp tục.",
                )

        presentations: dict[str, ContainerLoadPresentation] = {}
        for source_index, group in managed.items():
            runtime_id = runtime_by_source_index.get(source_index)
            if runtime_id is None:
                continue
            presentations[runtime_id] = ContainerLoadPresentation(
                status=group.status.value,
                message=(
                    (
                        f"{group_status_text(group)} – {group.bk_container_count} cont"
                        if group.status in {GroupStatus.ALLOCATED, GroupStatus.POSTED}
                        else group_status_text(group)
                    )
                ),
                session_id=str(group.id),
            )
        self.model.replace_lookup_presentations(presentations)
        self.model.set_contextual_warnings(warnings)

    def _lookup_action_clicked(self, proxy_index: QModelIndex) -> None:
        source_index = self.proxy_model.mapToSource(proxy_index)
        if not source_index.isValid():
            return
        source_row = source_index.row()
        runtime_id = self.model.runtime_id_at(source_row)
        presentation = self.model.lookup_presentation(runtime_id)
        if presentation.session_id:
            self.reconciliationOpenRequested.emit(int(presentation.session_id))
            return
        self._reconcile_source_row(source_row)

    @staticmethod
    def _missing_reconciliation_fields(row: ReviewRow) -> tuple[str, ...]:
        missing: list[str] = []
        if not isinstance(row.vessel_name, str) or not row.vessel_name.strip():
            missing.append("Tên tàu")
        if not isinstance(row.voyage_no, str) or not row.voyage_no.strip():
            missing.append("Số chuyến")
        if type(row.invoice_container_count) is not int or row.invoice_container_count <= 0:
            missing.append("SL cont HĐ")
        if row.container_count_basis not in {"EXPLICIT", "CALCULATED"}:
            missing.append("Căn cứ SL")
        if type(row.amount) is not int or row.amount < 0:
            missing.append("Số tiền")
        return tuple(missing)

    def _offer_edit_for_reconciliation(
        self,
        source_row: int,
        *,
        title: str,
        message: str,
    ) -> None:
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setWindowTitle(title)
        box.setText(message)
        edit_button = box.addButton("Sửa dòng", QMessageBox.ButtonRole.AcceptRole)
        box.addButton("Đóng", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() is edit_button:
            self._edit_source_row(source_row, focus_vessel=True)

    def _reconcile_source_row(self, source_row: int) -> bool:
        service = self._sea_freight_service
        if service is None:
            QMessageBox.warning(self, "Chưa thể đối soát", "Dịch vụ đối soát cước biển chưa được khởi tạo.")
            return False
        bk_path = str(_value(self._settings, "bk_workbook_path", default="") or "")
        row = self.model.row_at(source_row)
        missing = self._missing_reconciliation_fields(row)
        if missing:
            self._offer_edit_for_reconciliation(
                source_row,
                title="Chưa đủ thông tin đối soát",
                message=(
                    "Dòng cước biển còn thiếu hoặc chưa hợp lệ: "
                    + ", ".join(missing)
                    + ".\nHãy sửa dòng rồi bấm Đối soát lại."
                ),
            )
            return False
        data_row = DataRow.from_mapping(row.to_object())
        default_month: int | None = None
        default_year: int | None = None
        if isinstance(data_row.invoice_date, str):
            try:
                parsed = datetime.fromisoformat(data_row.invoice_date)
                default_month, default_year = parsed.month, parsed.year
            except ValueError:
                pass
        try:
            sheet_names = service.sheet_names(bk_path)
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Không đọc được file BK",
                f"Không thể đọc danh sách sheet đối soát.\n\n{exc}",
            )
            return False
        if not sheet_names:
            QMessageBox.warning(
                self,
                "Không có sheet đối soát",
                "File BK không có sheet tháng hợp lệ theo định dạng TMM YY.",
            )
            return False
        period_dialog = ReconciliationPeriodDialog(
            sheet_names=sheet_names,
            month=default_month,
            year=default_year,
            parent=self,
        )
        if period_dialog.exec() != QDialog.DialogCode.Accepted:
            return False
        month, year = period_dialog.period()
        managed_indices: set[int] = set()
        if self._batch_id is not None:
            try:
                managed_indices = set(
                    service.repository.managed_source_indices(int(self._batch_id))
                )
            except Exception:
                managed_indices = set()
        vessel_key = normalize_match_key(data_row.vessel_name)
        voyage_key = normalize_match_key(data_row.voyage_no)
        compatible = [
            (index, candidate)
            for index in range(self.model.rowCount())
            if index not in managed_indices
            and (candidate := self.model.row_at(index)).fee == "CB"
            and candidate.cont in (None, "")
            and not self._missing_reconciliation_fields(candidate)
            and normalize_match_key(candidate.vessel_name) == vessel_key
            and normalize_match_key(candidate.voyage_no) == voyage_key
        ]
        selection_dialog = SeaFreightContributionSelectionDialog(
            compatible, self
        )
        if selection_dialog.exec() != QDialog.DialogCode.Accepted:
            return False
        selected_indices = selection_dialog.selected_source_indices
        if not selected_indices:
            QMessageBox.warning(self, "Chưa chọn hóa đơn", "Hãy chọn ít nhất một hóa đơn.")
            return False
        outcome: Any | None = None
        try:
            open_many = getattr(service, "open_or_create_many", None)
            if callable(open_many):
                outcome = open_many(
                    [
                        (
                            self._source_index_by_runtime.get(
                                self.model.runtime_id_at(index), index
                            ),
                            DataRow.from_mapping(self.model.row_at(index).to_object()),
                        )
                        for index in selected_indices
                    ],
                    bk_path=bk_path,
                    month=month,
                    year=year,
                    source_batch_id=(
                        int(self._batch_id) if self._batch_id is not None else None
                    ),
                    source_sha256=str(
                        _value(self._metadata, "sha256", default="") or ""
                    ),
                )
                group = getattr(outcome, "group", outcome)
            else:
                group = service.open_or_create(
                    data_row,
                    bk_path=bk_path, month=month, year=year,
                    source_batch_id=int(self._batch_id) if self._batch_id is not None else None,
                    source_item_index=self._source_index_by_runtime.get(row.runtime_id, source_row),
                    source_sha256=str(_value(self._metadata, "sha256", default="") or ""),
                )
                selected_indices = [source_row]
        except VesselVoyageNotFoundError as exc:
            if exc.sheet_missing:
                QMessageBox.warning(self, "Không tìm thấy sheet BK", str(exc))
            else:
                dialog = VesselVoyageNotFoundDialog(exc, self)
                if dialog.exec() == VesselVoyageNotFoundDialog.EDIT_RESULT:
                    self._edit_source_row(source_row, focus_vessel=True)
            return False
        except Exception as exc:
            QMessageBox.warning(self, "Chưa thể đối soát số cont", str(exc))
            return False

        if bool(getattr(outcome, "requires_revision", False)):
            QMessageBox.information(
                self,
                "Hồ sơ đã hoàn tất",
                "Tàu/chuyến này đã có hồ sơ hoàn tất. Hãy mở hồ sơ, bấm "
                "Đối soát lại rồi thêm HĐ mới bằng Trợ lý ảo hoặc thêm tay.",
            )

        for index in selected_indices:
            selected_row = self.model.row_at(index)
            self.model.set_lookup_presentation(
                selected_row.runtime_id,
                status=(
                    "REQUIRES_REVISION"
                    if getattr(outcome, "requires_revision", False)
                    else group.status.value
                ),
                message=group_status_text(group),
                session_id=str(group.id),
            )
        if not bool(getattr(outcome, "requires_revision", False)):
            self.reconciliationChanged.emit()
        self.reconciliationOpenRequested.emit(group.id)
        return True
