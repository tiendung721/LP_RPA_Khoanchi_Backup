"""Cửa sổ thống nhất để bổ sung HĐ, kiểm tra cont BK và xác nhận phân bổ."""

from __future__ import annotations

import re
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QCloseEvent, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialogButtonBox,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QHeaderView,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from app.sea_freight.contracts import GroupStatus, group_status_text

from .app_dialog import AppDialog


def _money(value: object) -> str:
    try:
        return f"{int(value):,}".replace(",", ".")
    except (TypeError, ValueError):
        return str(value or "")


def _parse_int(value: str) -> int:
    text = value.strip().replace(".", "").replace(",", "").replace(" ", "")
    return int(text)


class ReconciliationPeriodDialog(AppDialog):
    def __init__(
        self,
        *,
        sheet_names: Any = (),
        month: int | None = None,
        year: int | None = None,
        parent: Any | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("reconciliationPeriodDialog")
        self.setWindowTitle("Chọn tháng đối soát")
        self.resize(800, 520)
        self._candidates = self._period_candidates(sheet_names)

        layout = QVBoxLayout(self)
        note = QLabel(
            "Hãy chọn một sheet để tiếp tục; chỉ sheet này được dùng để đối soát."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        self.table = QTableWidget(0, 3)
        self.table.setObjectName("reconciliationPeriodTable")
        self.table.setHorizontalHeaderLabels(("Sheet", "Tháng", "Năm"))
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table, 1)

        for sheet_name, candidate_month, candidate_year in self._candidates:
            row = self.table.rowCount()
            self.table.insertRow(row)
            sheet_item = QTableWidgetItem(sheet_name)
            sheet_item.setData(
                Qt.ItemDataRole.UserRole,
                (sheet_name, candidate_month, candidate_year),
            )
            self.table.setItem(row, 0, sheet_item)
            self.table.setItem(row, 1, QTableWidgetItem(str(candidate_month)))
            self.table.setItem(row, 2, QTableWidgetItem(str(candidate_year)))

        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Tiếp tục")
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Hủy")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self.table.itemSelectionChanged.connect(self._update_action)
        self.table.itemDoubleClicked.connect(lambda _item: self.accept())
        self._update_action()

        # Giữ tháng hóa đơn trong vùng nhìn thấy nếu có, nhưng không tự chọn thay user.
        if month is not None and year is not None:
            for row, (_sheet, candidate_month, candidate_year) in enumerate(
                self._candidates
            ):
                if (candidate_month, candidate_year) == (month, year):
                    self.table.scrollToItem(self.table.item(row, 0))
                    break

    @staticmethod
    def _period_candidates(sheet_names: Any) -> list[tuple[str, int, int]]:
        candidates: list[tuple[str, int, int]] = []
        for raw_name in sheet_names or ():
            name = str(raw_name).strip()
            match = re.fullmatch(r"T(\d{2})\s+(\d{2})", name, re.IGNORECASE)
            if match is None:
                continue
            month = int(match.group(1))
            year = 2000 + int(match.group(2))
            if 1 <= month <= 12:
                candidates.append((name, month, year))
        return sorted(candidates, key=lambda value: (value[2], value[1], value[0]), reverse=True)

    def _update_action(self) -> None:
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(
            self.table.currentRow() >= 0
        )

    @property
    def selected_sheet_name(self) -> str | None:
        row = self.table.currentRow()
        item = self.table.item(row, 0) if row >= 0 else None
        candidate = item.data(Qt.ItemDataRole.UserRole) if item is not None else None
        return str(candidate[0]) if candidate is not None else None

    def period(self) -> tuple[int, int]:
        row = self.table.currentRow()
        item = self.table.item(row, 0) if row >= 0 else None
        candidate = item.data(Qt.ItemDataRole.UserRole) if item is not None else None
        if candidate is None:
            raise ValueError("Chưa chọn sheet đối soát.")
        return int(candidate[1]), int(candidate[2])


class SeaFreightReconciliationDialog(AppDialog):
    changed = Signal()
    confirmed = Signal(object)

    SOURCE_LABELS = {
        "INITIAL": "HĐ ban đầu",
        "SUPPLEMENT": "HĐ bổ sung",
        "MANUAL": "Nhập tay",
    }
    INVOICE_HEADERS = (
        "Số HĐ", "Ngày HĐ", "B/L", "Số cont", "Số tiền", "Bên vận tải", "Nguồn dữ liệu"
    )
    CONTAINER_HEADERS = ("Số cont", "BK tháng", "Dòng BK", "SQT", "Ngày đi")
    RESULT_HEADERS = ("Số cont", "Số HĐ", "Tàu/chuyến", "Bên vận tải", "Tiền cước")

    def __init__(
        self,
        service: Any,
        *,
        batch_service: Any,
        open_assistant: Any,
        group_id: int,
        parent: Any | None = None,
    ) -> None:
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, False)
        self.service = service
        self.batch_service = batch_service
        self.open_assistant = open_assistant
        self.group_id = group_id
        self._loading = False
        self._dirty = False
        self._built_preview: list[Any] = []
        self.setWindowTitle("Đối soát số cont")
        self.setMinimumSize(1050, 720)
        self.resize(1280, 820)
        self._build_ui()
        self._connect_signals()
        self.load_group(group_id)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)

        heading = QHBoxLayout()
        title = QLabel("Đối soát số cont")
        title.setStyleSheet("font-size: 18pt; font-weight: 700;")
        heading.addWidget(title)
        heading.addStretch(1)
        heading.addWidget(QLabel("Hồ sơ đang mở:"))
        self.group_value = QLineEdit()
        self.group_value.setObjectName("currentReconciliationProfile")
        self.group_value.setReadOnly(True)
        self.group_value.setMinimumWidth(420)
        self.group_value.setToolTip(
            "Hồ sơ được cố định theo dòng đã chọn ở màn hình kiểm tra dữ liệu bóc tách."
        )
        heading.addWidget(self.group_value)
        root.addLayout(heading)

        common = QGroupBox("Thông tin chung")
        common_layout = QGridLayout(common)
        self.period_value = QLabel()
        self.vessel_name_edit = QLineEdit()
        self.voyage_edit = QLineEdit()
        self.vessel_preview = QLabel("—")
        self.vessel_preview.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        common_layout.addWidget(QLabel("Tháng đối soát:"), 0, 0)
        common_layout.addWidget(self.period_value, 0, 1)
        common_layout.addWidget(QLabel("Tàu/chuyến dùng để đối soát:"), 0, 2)
        common_layout.addWidget(self.vessel_preview, 0, 3)
        common_layout.addWidget(QLabel("Tên tàu:"), 1, 0)
        common_layout.addWidget(self.vessel_name_edit, 1, 1)
        common_layout.addWidget(QLabel("Số chuyến:"), 1, 2)
        common_layout.addWidget(self.voyage_edit, 1, 3)
        self.stats: dict[str, QLabel] = {}
        for column, (key, label) in enumerate(
            (("invoices", "Số HĐ"), ("invoice_containers", "Cont trên HĐ"),
             ("bk_containers", "Cont trong BK"), ("difference", "Chênh lệch"),
             ("amount", "Tổng tiền"))
        ):
            frame = QFrame()
            frame.setProperty("card", True)
            box = QVBoxLayout(frame)
            box.setContentsMargins(8, 4, 8, 4)
            caption = QLabel(label)
            caption.setProperty("muted", True)
            value = QLabel("0")
            value.setStyleSheet("font-weight: 700;")
            box.addWidget(caption)
            box.addWidget(value)
            common_layout.addWidget(frame, 2, column)
            self.stats[key] = value
        self.status_label = QLabel()
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet(
            "font-weight: 700; color: #92400E; background: #FFF7D6; padding: 7px; border-radius: 5px;"
        )
        common_layout.addWidget(self.status_label, 3, 0, 1, 5)
        root.addWidget(common)

        invoices_box = QGroupBox("Danh sách HĐ trong hồ sơ")
        invoices_layout = QVBoxLayout(invoices_box)
        self.invoice_table = QTableWidget(0, len(self.INVOICE_HEADERS))
        self.invoice_table.setHorizontalHeaderLabels(self.INVOICE_HEADERS)
        self.invoice_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.invoice_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.invoice_table.horizontalHeader().setStretchLastSection(True)
        self.invoice_table.setColumnWidth(0, 130)
        self.invoice_table.setColumnWidth(1, 105)
        self.invoice_table.setColumnWidth(2, 150)
        self.invoice_table.setColumnWidth(3, 90)
        self.invoice_table.setColumnWidth(4, 150)
        self.invoice_table.setColumnWidth(5, 190)
        invoices_layout.addWidget(self.invoice_table)
        root.addWidget(invoices_box, 2)

        lower = QHBoxLayout()
        containers_box = QGroupBox("Cont tìm được trong BK")
        containers_layout = QVBoxLayout(containers_box)
        self.container_table = QTableWidget(0, len(self.CONTAINER_HEADERS))
        self.container_table.setHorizontalHeaderLabels(self.CONTAINER_HEADERS)
        self.container_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.container_table.horizontalHeader().setStretchLastSection(True)
        containers_layout.addWidget(self.container_table)
        lower.addWidget(containers_box, 1)

        preview_box = QGroupBox("Kết quả dự kiến")
        preview_layout = QVBoxLayout(preview_box)
        self.result_table = QTableWidget(0, len(self.RESULT_HEADERS))
        self.result_table.setHorizontalHeaderLabels(self.RESULT_HEADERS)
        self.result_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.result_table.horizontalHeader().setStretchLastSection(True)
        preview_layout.addWidget(self.result_table)
        lower.addWidget(preview_box, 2)
        root.addLayout(lower, 2)

        actions = QHBoxLayout()
        self.assistant_button = QPushButton("Mở Trợ lý bóc tách")
        self.add_button = QPushButton("Thêm dòng HĐ")
        self.delete_button = QPushButton("Xóa dòng")
        actions.addWidget(self.assistant_button)
        actions.addWidget(self.add_button)
        actions.addWidget(self.delete_button)
        actions.addStretch(1)
        self.save_button = QPushButton("Lưu")
        self.rerun_button = QPushButton("Đối soát lại")
        self.confirm_button = QPushButton("Xác nhận")
        self.confirm_button.setProperty("primary", True)
        self.close_button = QPushButton("Đóng")
        actions.addWidget(self.save_button)
        actions.addWidget(self.rerun_button)
        actions.addWidget(self.confirm_button)
        actions.addWidget(self.close_button)
        root.addLayout(actions)

    def _connect_signals(self) -> None:
        self.vessel_name_edit.textChanged.connect(self._vessel_fields_changed)
        self.voyage_edit.textChanged.connect(self._vessel_fields_changed)
        self.invoice_table.itemChanged.connect(self._invoice_changed)
        self.assistant_button.clicked.connect(self._open_assistant)
        self.add_button.clicked.connect(self.add_invoice)
        self.delete_button.clicked.connect(self.delete_invoice)
        self.save_button.clicked.connect(self.save)
        self.rerun_button.clicked.connect(self.create_revision)
        self.confirm_button.clicked.connect(self.confirm)
        self.close_button.clicked.connect(self.close)

    def load_group(self, group_id: int) -> None:
        group = self.service.repository.get_group(group_id)
        if group is None:
            QMessageBox.warning(self, "Không mở được hồ sơ", "Hồ sơ không còn tồn tại.")
            return
        self.group_id = group_id
        self._loading = True
        try:
            self._show_current_group(group)
            self.period_value.setText(
                f"{group.reconciliation_month:02d}/{group.reconciliation_year}"
                if group.reconciliation_month and group.reconciliation_year
                else group.bk_sheet
            )
            contributions = self.service.repository.list_contributions(group_id)
            first = contributions[0] if contributions else None
            self.vessel_name_edit.setText(first.vessel_name if first else "")
            self.voyage_edit.setText(first.voyage_no if first else "")
            self._update_vessel_preview()
            self._load_invoices(contributions)
            self._load_containers(self.service.repository.list_container_rows(group_id))
            self._refresh_summary(group)
            locked = (
                not group.is_current
                or group.status in {GroupStatus.ALLOCATED, GroupStatus.POSTED, GroupStatus.CANCELLED}
            )
            self._set_editable(not locked)
            self.rerun_button.setEnabled(
                group.is_current and group.status in {GroupStatus.ALLOCATED, GroupStatus.POSTED}
            )
        finally:
            self._loading = False
            self._dirty = False

    def _show_current_group(self, group: Any) -> None:
        label = (
            f"{group.vessel_voyage_raw} – Lần {group.revision_no} – "
            f"{group_status_text(group)}"
            + (" – Phiên cũ" if not group.is_current else "")
        )
        self.group_value.setText(label)
        self.group_value.setToolTip(
            "Hồ sơ được cố định theo dòng đã chọn ở màn hình kiểm tra dữ liệu bóc tách. "
            f"ID hồ sơ: {group.id}."
        )

    def _load_invoices(self, contributions: list[Any]) -> None:
        self.invoice_table.setRowCount(len(contributions))
        for row_index, item in enumerate(contributions):
            values = (
                item.invoice_no or "", item.invoice_date or "", item.bl or "",
                str(item.invoice_container_count), _money(item.amount), item.carrier or "",
                self.SOURCE_LABELS.get(item.source_kind, "HĐ ban đầu"),
            )
            for column, value in enumerate(values):
                cell = QTableWidgetItem(value)
                if column == 0:
                    cell.setData(Qt.ItemDataRole.UserRole, item.id)
                    cell.setData(Qt.ItemDataRole.UserRole + 1, item.source_kind)
                    cell.setData(Qt.ItemDataRole.UserRole + 2, item.source_batch_id)
                    cell.setData(Qt.ItemDataRole.UserRole + 3, item.source_item_index)
                    cell.setData(Qt.ItemDataRole.UserRole + 4, item.source_sha256)
                    cell.setData(Qt.ItemDataRole.UserRole + 5, item.source_document_id)
                    cell.setData(Qt.ItemDataRole.UserRole + 6, item.source_document_name)
                if column == 6:
                    cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if item.source_kind == "SUPPLEMENT":
                    cell.setBackground(QColor("#E8F2FF"))
                self.invoice_table.setItem(row_index, column, cell)

    def _load_containers(self, rows: list[dict[str, Any]]) -> None:
        self.container_table.setRowCount(len(rows))
        for row_index, item in enumerate(rows):
            values = (
                item.get("container"), item.get("source_sheet"), item.get("source_row"),
                item.get("source_sqt"), item.get("departure_date"),
            )
            for column, value in enumerate(values):
                self.container_table.setItem(row_index, column, QTableWidgetItem(str(value or "")))
        self.container_table.resizeColumnsToContents()

    def _refresh_summary(self, group: Any) -> None:
        invoice_count = self.invoice_table.rowCount()
        self.stats["invoices"].setText(str(invoice_count))
        self.stats["invoice_containers"].setText(str(group.received_container_count))
        self.stats["bk_containers"].setText(str(group.bk_container_count))
        difference = group.received_container_count - group.bk_container_count
        self.stats["difference"].setText(f"{difference:+d}" if difference else "0")
        self.stats["amount"].setText(_money(group.total_amount) + " đ")
        status = group_status_text(group)
        warnings: list[str] = []
        if group.bk_duplicate_container_count:
            warnings.append(f"BK có {group.bk_duplicate_container_count} dòng cont trùng đã bỏ qua")
        if group.bk_invalid_container_count:
            warnings.append(f"BK có {group.bk_invalid_container_count} số cont sai đã bỏ qua")
        self.status_label.setText(" • ".join((status, *warnings)))
        self._built_preview = []
        if group.status in {
            GroupStatus.READY,
            GroupStatus.ALLOCATED,
            GroupStatus.POSTED,
        }:
            try:
                self._built_preview = self.service.allocation_rows(group.id)
            except Exception:
                self._built_preview = []
        self._load_preview(self._built_preview)
        self.confirm_button.setEnabled(group.status is GroupStatus.READY and bool(self._built_preview))

    def _load_preview(self, rows: list[Any]) -> None:
        self.result_table.setRowCount(len(rows))
        for row_index, item in enumerate(rows):
            values = (item.cont, item.invoice_no, item.vessel_voyage_raw, item.carrier, _money(item.amount))
            for column, value in enumerate(values):
                self.result_table.setItem(row_index, column, QTableWidgetItem(str(value or "")))
        self.result_table.resizeColumnsToContents()

    def _set_editable(self, editable: bool) -> None:
        for widget in (
            self.vessel_name_edit, self.voyage_edit,
            self.invoice_table, self.assistant_button, self.add_button,
            self.delete_button, self.save_button,
        ):
            widget.setEnabled(editable)

    def _mark_dirty(self, *_args: Any) -> None:
        if not self._loading:
            self._dirty = True
            self.confirm_button.setEnabled(False)

    def _update_vessel_preview(self) -> None:
        vessel = " ".join(self.vessel_name_edit.text().strip().split())
        voyage = " ".join(self.voyage_edit.text().strip().split())
        self.vessel_preview.setText(
            " ".join(part for part in (vessel, voyage) if part) or "—"
        )

    def _vessel_fields_changed(self, *_args: Any) -> None:
        self._update_vessel_preview()
        self._mark_dirty()

    def _invoice_changed(self, *_args: Any) -> None:
        self._mark_dirty()
        self._refresh_draft_summary()

    def _refresh_draft_summary(self) -> None:
        if self._loading:
            return
        try:
            payload = self._invoice_payload()
            invoice_containers = sum(int(item["invoice_container_count"]) for item in payload)
            total = sum(int(item["amount"]) for item in payload)
        except (AttributeError, TypeError, ValueError):
            self.status_label.setText("Cần sửa dữ liệu HĐ • Chưa lưu")
            self.result_table.setRowCount(0)
            return
        bk_count = self.container_table.rowCount()
        difference = invoice_containers - bk_count
        self.stats["invoices"].setText(str(len(payload)))
        self.stats["invoice_containers"].setText(str(invoice_containers))
        self.stats["difference"].setText(f"{difference:+d}" if difference else "0")
        self.stats["amount"].setText(_money(total) + " đ")
        if bk_count == 0:
            text = "Chưa tìm thấy dữ liệu BK"
        elif difference < 0:
            text = f"Thiếu {-difference} cont"
        elif difference > 0:
            text = f"Thừa {difference} cont"
        else:
            text = f"Đủ {invoice_containers}/{bk_count} cont"
        self.status_label.setText(text + " • Chưa lưu")
        self.result_table.setRowCount(0)

    def add_invoice(self) -> None:
        row = self.invoice_table.rowCount()
        self.invoice_table.insertRow(row)
        for column in range(len(self.INVOICE_HEADERS)):
            cell = QTableWidgetItem("Nhập tay" if column == 6 else "")
            if column == 0:
                cell.setData(Qt.ItemDataRole.UserRole + 1, "MANUAL")
                cell.setData(Qt.ItemDataRole.UserRole + 5, "MANUAL")
                cell.setData(Qt.ItemDataRole.UserRole + 6, "Dòng thêm thủ công")
            if column == 6:
                cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.invoice_table.setItem(row, column, cell)
        self.invoice_table.selectRow(row)
        self.invoice_table.editItem(self.invoice_table.item(row, 0))
        self._mark_dirty()
        self._refresh_draft_summary()

    def delete_invoice(self) -> None:
        row = self.invoice_table.currentRow()
        if row < 0:
            return
        self.invoice_table.removeRow(row)
        self._mark_dirty()
        self._refresh_draft_summary()

    def _invoice_payload(self) -> list[dict[str, object]]:
        result: list[dict[str, object]] = []
        for row in range(self.invoice_table.rowCount()):
            first = self.invoice_table.item(row, 0)
            result.append(
                {
                    "id": first.data(Qt.ItemDataRole.UserRole),
                    "source_kind": first.data(Qt.ItemDataRole.UserRole + 1) or "MANUAL",
                    "source_batch_id": first.data(Qt.ItemDataRole.UserRole + 2),
                    "source_item_index": first.data(Qt.ItemDataRole.UserRole + 3) or 0,
                    "source_sha256": first.data(Qt.ItemDataRole.UserRole + 4),
                    "source_document_id": first.data(Qt.ItemDataRole.UserRole + 5) or "MANUAL",
                    "source_document_name": first.data(Qt.ItemDataRole.UserRole + 6) or "Dòng thêm thủ công",
                    "invoice_no": first.text().strip(),
                    "invoice_date": self.invoice_table.item(row, 1).text().strip(),
                    "bl": self.invoice_table.item(row, 2).text().strip(),
                    "invoice_container_count": _parse_int(self.invoice_table.item(row, 3).text()),
                    "amount": _parse_int(self.invoice_table.item(row, 4).text()),
                    "carrier": self.invoice_table.item(row, 5).text().strip(),
                }
            )
        return result

    def save(self, *, show_success: bool = True) -> bool:
        try:
            group = self.service.save_group(
                self.group_id,
                vessel_voyage_raw=self.vessel_preview.text().strip(),
                vessel_name=self.vessel_name_edit.text().strip(),
                voyage_no=self.voyage_edit.text().strip(),
                invoices=self._invoice_payload(),
            )
        except Exception as exc:
            QMessageBox.warning(self, "Chưa lưu được", str(exc))
            return False
        self.load_group(group.id)
        self.changed.emit()
        if show_success:
            self.status_label.setText(group_status_text(group) + " • Đã lưu")
        return True

    def confirm(self) -> None:
        if self._dirty and not self.save(show_success=False):
            return
        try:
            rows = self.service.prepare_allocation(self.group_id)
            total = sum(int(row.amount or 0) for row in rows)
        except Exception as exc:
            QMessageBox.warning(self, "Chưa thể xác nhận", str(exc))
            self.load_group(self.group_id)
            return
        answer = QMessageBox.question(
            self,
            "Xác nhận kết quả",
            f"Tạo {len(rows)} dòng cước theo {len(rows)} cont với tổng tiền {_money(total)} đồng?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            review = self.batch_service.create_reconciliation_batch(self.group_id, rows)
            group = self.service.repository.get_group(self.group_id)
            if group is not None and group.status is GroupStatus.READY:
                self.service.repository.mark_allocated(self.group_id, review.metadata.id)
        except Exception as exc:
            QMessageBox.critical(self, "Không tạo được kết quả", str(exc))
            return
        self.load_group(self.group_id)
        self.changed.emit()
        self.confirmed.emit(review)

    def create_revision(self) -> None:
        group = self.service.repository.get_group(self.group_id)
        if group is None:
            return
        detail = (
            "Kết quả hiện tại đã được ghi vào BK. Phiên mới sẽ không tự ghi đè; "
            "mọi ô có dữ liệu khác phải được xác nhận thủ công.\n\n"
            if group.status is GroupStatus.POSTED
            else "Kết quả hiện tại sẽ được lưu lịch sử và batch cũ sẽ được lưu trữ.\n\n"
        )
        answer = QMessageBox.question(
            self,
            "Đối soát lại",
            detail + f"Tạo lần đối soát {group.revision_no + 1}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            created = self.service.create_revision(group.id)
        except Exception as exc:
            QMessageBox.warning(self, "Không thể đối soát lại", str(exc))
            return
        self.load_group(created.id)
        self.changed.emit()

    def _open_assistant(self) -> None:
        if callable(self.open_assistant):
            self.open_assistant()

    def supplement_received(self, result: Any) -> None:
        self.load_group(self.group_id)
        new_keys = {str(value).strip().casefold() for value in result.added_invoices}
        self._loading = True
        try:
            for row in range(self.invoice_table.rowCount()):
                invoice = self.invoice_table.item(row, 0).text().strip().casefold()
                if invoice not in new_keys:
                    continue
                self.invoice_table.item(row, 6).setText("Mới nhận")
                for column in range(self.invoice_table.columnCount()):
                    self.invoice_table.item(row, column).setBackground(QColor("#DCEEFF"))
        finally:
            self._loading = False
            self._dirty = False
        parts: list[str] = []
        if result.added_count:
            parts.append(f"Đã nhận thêm {result.added_count} HĐ")
        if result.duplicate_invoices:
            parts.append("HĐ đã có: " + ", ".join(result.duplicate_invoices))
        if not result.added_count and not result.duplicate_invoices:
            parts.append("File mới không có HĐ thuộc tàu/chuyến này")
        if result.errors:
            parts.append("; ".join(result.errors))
        self.status_label.setText(" • ".join(parts))
        self.raise_()
        self.activateWindow()

    def _resolve_unsaved(self) -> bool:
        if not self._dirty:
            return True
        box = QMessageBox(self)
        box.setWindowTitle("Thay đổi chưa lưu")
        box.setText("Hồ sơ có thay đổi chưa lưu. Bạn có muốn lưu trước khi đóng?")
        save = box.addButton("Lưu và đóng", QMessageBox.ButtonRole.AcceptRole)
        discard = box.addButton("Đóng không lưu", QMessageBox.ButtonRole.DestructiveRole)
        box.addButton("Tiếp tục chỉnh sửa", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() is save:
            return self.save(show_success=False)
        if box.clickedButton() is discard:
            self._dirty = False
            return True
        return False

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802
        if self._resolve_unsaved():
            event.accept()
        else:
            event.ignore()


# Tên cũ được giữ làm alias để code tích hợp ngoài dự án không vỡ ngay.
SeaFreightCenterDialog = SeaFreightReconciliationDialog

__all__ = [
    "ReconciliationPeriodDialog",
    "SeaFreightCenterDialog",
    "SeaFreightReconciliationDialog",
]
