"""Hộp thoại thêm/sửa một dòng dữ liệu bóc tách."""

from __future__ import annotations

import re
from collections.abc import Callable
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (
    QComboBox,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.services.carrier_policy import carrier_is_managed_by_daily_sync

from .app_dialog import AppDialog
from .review_table_model import (
    FEE_CATALOG,
    RULE_CATALOG,
    ReviewRow,
    RowValidation,
    coerce_review_row,
    validate_row,
)

try:
    from app.services.validation_service import (
        normalize_bl as _core_normalize_bl,
        normalize_container as _core_normalize_container,
        normalize_optional_text as _core_normalize_optional_text,
        parse_amount as _core_parse_amount,
    )
except ImportError:
    _core_normalize_bl = None
    _core_normalize_container = None
    _core_normalize_optional_text = None
    _core_parse_amount = None


def normalize_container(value: str) -> str | None:
    """Chuẩn hóa container mà không suy đoán ký tự OCR."""

    if _core_normalize_container is not None:
        return _core_normalize_container(value)
    compact = re.sub(r"[\s\-_–—]+", "", value.strip()).upper()
    return compact or None


def normalize_bl(value: str) -> str | None:
    """Viết hoa, trim và thu gọn khoảng trắng nhưng giữ dấu có ý nghĩa."""

    if _core_normalize_bl is not None:
        return _core_normalize_bl(value)
    normalized = re.sub(r"\s+", " ", value.strip()).upper()
    return normalized or None


def normalize_optional_text(value: str) -> str | None:
    if _core_normalize_optional_text is not None:
        return _core_normalize_optional_text(value)
    normalized = re.sub(r"\s+", " ", value.strip())
    return normalized or None


def parse_amount(text: str, *, allow_negative: bool = False) -> int | None:
    """Đọc số nguyên 64 bit từ cách gõ thân thiện có phân cách hàng nghìn."""

    if _core_parse_amount is not None:
        return _core_parse_amount(text, allow_negative=allow_negative)
    value = text.strip()
    if not value:
        return None
    sign = -1 if value.startswith("-") else 1
    unsigned = value[1:].strip() if sign < 0 else value
    if sign < 0 and not allow_negative:
        raise ValueError("Số tiền không được âm.")
    if not re.fullmatch(r"\d+", unsigned):
        if not re.fullmatch(r"\d{1,3}(?:[., ]\d{3})+", unsigned):
            raise ValueError(
                "Số tiền chỉ được chứa chữ số và dấu phân cách hàng nghìn hợp lệ."
            )
    compact = re.sub(r"[., ]", "", unsigned)
    amount = sign * int(compact)
    if abs(amount) > 9_223_372_036_854_775_807:
        raise ValueError("Số tiền vượt giới hạn số nguyên 64 bit.")
    return amount


def format_amount(value: int | None) -> str:
    return "" if value is None else f"{value:,}".replace(",", ".")


class EditRowDialog(AppDialog):
    """Dialog chỉnh sửa có chuẩn hóa, parser 64-bit và validation realtime."""

    rowAccepted = Signal(object)

    def __init__(
        self,
        row: Any | None = None,
        parent: QWidget | None = None,
        *,
        validator: Callable[[ReviewRow], Any] | Any | None = None,
        title: str | None = None,
        allow_negative: bool = False,
        source_document_id: str = "MANUAL",
        source_document_name: str = "Dòng thêm thủ công",
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title or ("Thêm dòng dữ liệu" if row is None else "Sửa dòng dữ liệu"))
        self.setModal(True)
        self.resize(680, 760)
        self._editing = row is not None
        self._original_rule: Any = None
        self._original_vessel_voyage_raw: Any = None
        self._validator = validator
        self._allow_negative = bool(allow_negative)
        self._result_row: ReviewRow | None = None
        self._last_validation = RowValidation()
        self._last_fee: object = None
        self._source_document_id = source_document_id
        self._source_document_name = source_document_name
        self._build_ui()
        self._connect_signals()
        self.set_row(
            row
            if row is not None
            else ReviewRow(
                source_document_id=source_document_id,
                source_document_name=source_document_name,
            )
        )

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 18)
        root.setSpacing(14)

        heading = QLabel(self.windowTitle())
        heading.setObjectName("pageTitle")
        root.addWidget(heading)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(12)

        self.source_document_label = QLabel()
        self.source_document_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.source_document_label.setWordWrap(True)
        form.addRow("Chứng từ nguồn:", self.source_document_label)

        self.container_edit = QLineEdit()
        self.container_edit.setObjectName("containerEdit")
        self.container_edit.setPlaceholderText("Ví dụ: DRYU3026167")
        self.container_edit.setMaxLength(40)
        self.container_edit.setClearButtonEnabled(True)
        self.container_hint = QLabel("Mẫu thường gặp: 4 chữ cái + 7 chữ số; sai mẫu chỉ là cảnh báo.")
        self.container_hint.setProperty("muted", True)
        container_box = QVBoxLayout()
        container_box.setContentsMargins(0, 0, 0, 0)
        container_box.setSpacing(3)
        container_box.addWidget(self.container_edit)
        container_box.addWidget(self.container_hint)
        container_widget = QWidget()
        container_widget.setLayout(container_box)
        form.addRow("Container:", container_widget)

        self.bl_edit = QLineEdit()
        self.bl_edit.setObjectName("blEdit")
        self.bl_edit.setPlaceholderText("Để trống nếu chưa xác định")
        self.bl_edit.setMaxLength(200)
        self.bl_edit.setClearButtonEnabled(True)
        form.addRow("Số B/L:", self.bl_edit)

        self.vessel_name_edit = QLineEdit()
        self.vessel_name_edit.setObjectName("vesselNameEdit")
        self.vessel_name_edit.setPlaceholderText("Ví dụ: PROSPER")
        self.vessel_name_edit.setClearButtonEnabled(True)
        form.addRow("Tên tàu:", self.vessel_name_edit)

        self.voyage_no_edit = QLineEdit()
        self.voyage_no_edit.setObjectName("voyageNoEdit")
        self.voyage_no_edit.setPlaceholderText("Ví dụ: 2625S")
        self.voyage_no_edit.setClearButtonEnabled(True)
        form.addRow("Số chuyến:", self.voyage_no_edit)

        self.vessel_voyage_preview = QLabel("—")
        self.vessel_voyage_preview.setObjectName("vesselVoyagePreview")
        self.vessel_voyage_preview.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        form.addRow("Tàu/chuyến dùng để đối soát:", self.vessel_voyage_preview)

        self.invoice_container_count_edit = QLineEdit()
        self.invoice_container_count_edit.setPlaceholderText("Số nguyên dương")
        form.addRow("SL cont HĐ:", self.invoice_container_count_edit)

        self.container_count_basis_combo = QComboBox()
        self.container_count_basis_combo.addItem("UNKNOWN – Chưa xác định", "UNKNOWN")
        self.container_count_basis_combo.addItem("EXPLICIT – Ghi rõ trên HĐ", "EXPLICIT")
        self.container_count_basis_combo.addItem("CALCULATED – Tính từ HĐ", "CALCULATED")
        form.addRow("Căn cứ SL:", self.container_count_basis_combo)
        self._sea_freight_widgets = (
            self.vessel_name_edit,
            self.voyage_no_edit,
            self.vessel_voyage_preview,
            self.invoice_container_count_edit,
            self.container_count_basis_combo,
        )
        self._form = form

        self.fee_combo = QComboBox()
        self.fee_combo.setObjectName("feeCombo")
        self.fee_combo.addItem("— Chọn loại cước —", None)
        for code, name in FEE_CATALOG.items():
            self.fee_combo.addItem(f"{code} – {name}", code)
        self.fee_combo.setMaxVisibleItems(14)
        form.addRow("Loại cước:", self.fee_combo)

        self.rule_combo = QComboBox()
        self.rule_combo.setObjectName("ruleCombo")
        for code, name in RULE_CATALOG.items():
            prefix = "null" if code is None else code
            self.rule_combo.addItem(f"{prefix} – {name}", code)
        form.addRow("Quy tắc tiền:", self.rule_combo)
        form.setRowVisible(self.rule_combo, not self._editing)

        self.invoice_no_edit = QLineEdit()
        self.invoice_no_edit.setObjectName("invoiceNoEdit")
        self.invoice_no_edit.setPlaceholderText("Để trống nếu chưa xác định")
        self.invoice_no_edit.setMaxLength(200)
        self.invoice_no_edit.setClearButtonEnabled(True)
        form.addRow("Số HĐ:", self.invoice_no_edit)

        self.invoice_date_edit = QLineEdit()
        self.invoice_date_edit.setPlaceholderText("YYYY-MM-DD")
        self.invoice_date_edit.setClearButtonEnabled(True)
        form.addRow("Ngày HĐ:", self.invoice_date_edit)

        self.carrier_edit = QLineEdit()
        self.carrier_edit.setObjectName("carrierEdit")
        self.carrier_edit.setPlaceholderText("Để trống nếu chưa xác định")
        self.carrier_edit.setMaxLength(300)
        self.carrier_edit.setClearButtonEnabled(True)
        form.addRow("Bên vận tải:", self.carrier_edit)

        amount_box = QVBoxLayout()
        amount_box.setContentsMargins(0, 0, 0, 0)
        amount_box.setSpacing(5)
        self.amount_edit = QLineEdit()
        self.amount_edit.setObjectName("amountEdit")
        self.amount_edit.setPlaceholderText("Ví dụ: 13.554.000")
        self.amount_edit.setMaxLength(30)
        self.amount_edit.setAlignment(Qt.AlignmentFlag.AlignRight)
        amount_box.addWidget(self.amount_edit)
        amount_hint = QLabel(
            "Số nguyên VND có dấu; số âm là khoản điều chỉnh giảm."
            if self._allow_negative
            else "Lưu dưới dạng số nguyên VND; không nhập ký hiệu tiền tệ hoặc số âm."
        )
        amount_hint.setProperty("muted", True)
        amount_box.addWidget(amount_hint)
        amount_widget = QWidget()
        amount_widget.setLayout(amount_box)
        form.addRow("Số tiền:", amount_widget)
        root.addLayout(form)

        validation_label = QLabel("Kiểm tra realtime")
        validation_label.setStyleSheet("font-weight: 600;")
        root.addWidget(validation_label)
        self.validation_view = QPlainTextEdit()
        self.validation_view.setObjectName("validationView")
        self.validation_view.setReadOnly(True)
        self.validation_view.setMaximumHeight(112)
        self.validation_view.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        root.addWidget(self.validation_view)

        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        self.save_button = self.button_box.button(QDialogButtonBox.StandardButton.Save)
        self.cancel_button = self.button_box.button(QDialogButtonBox.StandardButton.Cancel)
        self.save_button.setText("Lưu dòng")
        self.save_button.setObjectName("saveRowButton")
        self.save_button.setProperty("primary", True)
        self.cancel_button.setText("Hủy")
        root.addWidget(self.button_box)

    def _connect_signals(self) -> None:
        self.container_edit.textChanged.connect(self._validate_realtime)
        self.bl_edit.textChanged.connect(self._validate_realtime)
        self.vessel_name_edit.textChanged.connect(self._vessel_changed)
        self.voyage_no_edit.textChanged.connect(self._vessel_changed)
        self.invoice_container_count_edit.textChanged.connect(self._validate_realtime)
        self.container_count_basis_combo.currentIndexChanged.connect(self._validate_realtime)
        self.fee_combo.currentIndexChanged.connect(self._fee_changed)
        self.rule_combo.currentIndexChanged.connect(self._validate_realtime)
        self.invoice_no_edit.textChanged.connect(self._validate_realtime)
        self.invoice_date_edit.textChanged.connect(self._validate_realtime)
        self.carrier_edit.textChanged.connect(self._validate_realtime)
        self.amount_edit.textChanged.connect(self._validate_realtime)
        self.amount_edit.editingFinished.connect(self._format_amount_on_finish)
        self.button_box.accepted.connect(self._accept_row)
        self.button_box.rejected.connect(self.reject)

    def _fee_changed(self, *_args: Any) -> None:
        current = self.fee_combo.currentData()
        if (
            carrier_is_managed_by_daily_sync(current)
            and not carrier_is_managed_by_daily_sync(self._last_fee)
        ):
            # Khi đổi một HĐ phí thành loại vận tải lấy từ Hàng ngày, carrier
            # cũ không được vô tình trở thành một ghi đè thủ công. Ô vẫn mở để
            # user nhập lại sau thao tác đổi mã nếu đó thực sự là chủ ý.
            self.carrier_edit.clear()
        self._last_fee = current
        self._update_sea_freight_visibility()
        self._validate_realtime()

    def _vessel_changed(self, *_args: Any) -> None:
        vessel = " ".join(self.vessel_name_edit.text().strip().split())
        voyage = " ".join(self.voyage_no_edit.text().strip().split())
        self.vessel_voyage_preview.setText(
            " ".join(part for part in (vessel, voyage) if part) or "—"
        )
        self._validate_realtime()

    def _update_sea_freight_visibility(self) -> None:
        visible = self.fee_combo.currentData() == "CB"
        for widget in self._sea_freight_widgets:
            self._form.setRowVisible(widget, visible)

    def focus_vessel_fields(self) -> None:
        self.vessel_name_edit.setFocus(Qt.FocusReason.OtherFocusReason)
        self.vessel_name_edit.selectAll()

    def set_row(self, row: Any) -> None:
        value = coerce_review_row(row)
        self._source_document_id = value.source_document_id
        self._source_document_name = value.source_document_name
        self.source_document_label.setText(
            f"{self._source_document_name} ({self._source_document_id})"
        )
        self._original_vessel_voyage_raw = value.vessel_voyage_raw
        self.container_edit.setText(value.cont if isinstance(value.cont, str) else "")
        self.bl_edit.setText(value.bl if isinstance(value.bl, str) else "")
        self.vessel_name_edit.setText(
            value.vessel_name if isinstance(value.vessel_name, str) else ""
        )
        self.voyage_no_edit.setText(
            value.voyage_no if isinstance(value.voyage_no, str) else ""
        )
        self.invoice_container_count_edit.setText(
            str(value.invoice_container_count)
            if type(value.invoice_container_count) is int
            else ""
        )
        self._select_combo_data(
            self.container_count_basis_combo, value.container_count_basis or "UNKNOWN"
        )
        self._select_combo_data(self.fee_combo, value.fee)
        self._original_rule = value.rule
        self._select_combo_data(self.rule_combo, value.rule)
        self.invoice_no_edit.setText(
            value.invoice_no if isinstance(value.invoice_no, str) else ""
        )
        self.invoice_date_edit.setText(
            value.invoice_date if isinstance(value.invoice_date, str) else ""
        )
        self.carrier_edit.setText(
            value.carrier if isinstance(value.carrier, str) else ""
        )
        valid_amount = (
            value.amount
            if isinstance(value.amount, int) and not isinstance(value.amount, bool)
            else None
        )
        self.amount_edit.setText(format_amount(valid_amount))
        self._vessel_changed()
        self._update_sea_freight_visibility()
        self._validate_realtime()

    @staticmethod
    def _select_combo_data(combo: QComboBox, data: Any) -> None:
        index = combo.findData(data)
        if index < 0 and data is not None:
            combo.insertItem(0, f"{data} – Mã không hợp lệ, hãy chọn lại", data)
            combo.setItemData(0, "Mã này không thuộc danh mục chính thức.", Qt.ItemDataRole.ToolTipRole)
            index = 0
        combo.setCurrentIndex(max(index, 0))

    def _format_amount_on_finish(self) -> None:
        try:
            self.amount_edit.setText(
                format_amount(
                    parse_amount(
                        self.amount_edit.text(), allow_negative=self._allow_negative
                    )
                )
            )
        except ValueError:
            return

    def _collect_row(self) -> tuple[ReviewRow, str | None]:
        amount_error: str | None = None
        amount: int | None = None
        try:
            amount = parse_amount(
                self.amount_edit.text(), allow_negative=self._allow_negative
            )
        except ValueError as exc:
            amount_error = str(exc)
        count_text = self.invoice_container_count_edit.text().strip()
        count: int | None = None
        if count_text:
            if count_text.isdigit() and int(count_text) > 0:
                count = int(count_text)
            else:
                amount_error = "SL cont HĐ phải là số nguyên dương."
        row = ReviewRow(
            cont=normalize_container(self.container_edit.text()),
            bl=normalize_bl(self.bl_edit.text()),
            fee=self.fee_combo.currentData(),
            rule=self._original_rule if self._editing else self.rule_combo.currentData(),
            amount=amount,
            invoice_no=normalize_optional_text(self.invoice_no_edit.text()),
            carrier=normalize_optional_text(self.carrier_edit.text()),
            vessel_voyage_raw=self._original_vessel_voyage_raw,
            vessel_name=normalize_optional_text(self.vessel_name_edit.text()),
            voyage_no=normalize_optional_text(self.voyage_no_edit.text()),
            invoice_container_count=count,
            container_count_basis=self.container_count_basis_combo.currentData(),
            invoice_date=normalize_optional_text(self.invoice_date_edit.text()),
            source_document_id=self._source_document_id,
            source_document_name=self._source_document_name,
        )
        return row, amount_error

    def _validate_realtime(self, *_args: Any) -> None:
        row, amount_error = self._collect_row()
        local = validate_row(row, allow_negative=self._allow_negative)
        errors = list(local.errors)
        warnings = list(local.warnings)
        if amount_error:
            errors.append(amount_error)

        external_errors, external_warnings = self._external_messages(row)
        errors.extend(message for message in external_errors if message not in errors)
        warnings.extend(message for message in external_warnings if message not in warnings)
        self._last_validation = RowValidation(tuple(errors), tuple(warnings))
        self.save_button.setEnabled(not errors)

        self.container_edit.setProperty(
            "invalid",
            any("Container phải" in message for message in errors),
        )
        self.fee_combo.setProperty(
            "invalid",
            any("cước" in message.casefold() for message in errors),
        )
        self.rule_combo.setProperty(
            "invalid",
            any(("quy tắc" in message.casefold() or "HD" in message) for message in errors),
        )
        self.amount_edit.setProperty(
            "invalid",
            bool(amount_error) or any("tiền" in message.casefold() for message in errors),
        )
        for widget in (
            self.container_edit,
            self.fee_combo,
            self.rule_combo,
            self.amount_edit,
        ):
            widget.style().unpolish(widget)
            widget.style().polish(widget)

        if errors:
            lines = ["LỖI – cần sửa trước khi lưu:"] + [f"• {message}" for message in errors]
            if warnings:
                lines += ["", "Cảnh báo:"] + [f"• {message}" for message in warnings]
            self.validation_view.setStyleSheet("color: #B42318; background: #FFF7F7;")
        elif warnings:
            lines = ["CẢNH BÁO – có thể lưu sau khi kiểm tra:"] + [
                f"• {message}" for message in warnings
            ]
            self.validation_view.setStyleSheet("color: #8A5A00; background: #FFFBEB;")
        else:
            lines = ["HỢP LỆ – không phát hiện lỗi hoặc cảnh báo."]
            self.validation_view.setStyleSheet("color: #15803D; background: #F0FDF4;")
        self.validation_view.setPlainText("\n".join(lines))

    def _external_messages(self, row: ReviewRow) -> tuple[list[str], list[str]]:
        if self._validator is None:
            return [], []
        try:
            if callable(self._validator):
                result = self._validator(row)
            else:
                method = getattr(self._validator, "validate_row", None)
                if not callable(method):
                    return [], []
                if self._allow_negative:
                    try:
                        result = method(row, allow_negative=True)
                    except TypeError:
                        result = method(row)
                else:
                    result = method(row)
        except Exception:
            return [], []

        errors = list(getattr(result, "errors", ()) or ())
        warnings = list(getattr(result, "warnings", ()) or ())
        issues = getattr(result, "issues", ()) or ()
        for issue in issues:
            message = str(getattr(issue, "message", issue))
            severity = str(getattr(issue, "severity", "")).casefold()
            if "error" in severity or "lỗi" in severity:
                errors.append(message)
            elif "warning" in severity or "cảnh báo" in severity:
                warnings.append(message)
        return [str(item) for item in errors], [str(item) for item in warnings]

    def _accept_row(self) -> None:
        row, amount_error = self._collect_row()
        self._validate_realtime()
        if amount_error or self._last_validation.errors:
            QMessageBox.warning(
                self,
                "Chưa thể lưu dòng",
                "Dòng vẫn còn lỗi chặn. Hãy sửa các mục được đánh dấu rồi thử lại.",
            )
            return
        self._result_row = row
        self.rowAccepted.emit(self.row_data())
        self.accept()

    def row_data(self) -> ReviewRow:
        """Trả về bản sao dòng đã lưu (hoặc dữ liệu hợp lệ đang nhập)."""

        if self._result_row is not None:
            return ReviewRow.from_mapping(self._result_row.to_object())
        row, error = self._collect_row()
        if error:
            raise ValueError(error)
        return row

    get_row = row_data

    def row_array(self) -> list[Any]:
        return self.row_data().as_array()

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802
        # Đóng bằng nút X tương đương Hủy; dialog chưa ghi gì vào model nguồn.
        event.accept()


# Tên tương thích dễ đoán cho lớp tích hợp.
RowEditDialog = EditRowDialog
