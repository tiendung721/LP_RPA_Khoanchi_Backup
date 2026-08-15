"""Trang cấu hình BAT, Output và hai workbook dùng cho Bước 3."""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from .feedback import LinearLoadingBar, set_button_loading


def _setting(source: Any, name: str, default: Any = "") -> Any:
    if source is None:
        return default
    if isinstance(source, Mapping):
        return source.get(name, default)
    return getattr(source, name, default)


class SettingsPage(QWidget):
    save_requested = Signal(object)
    check_requested = Signal(object)
    open_output_requested = Signal(object)

    def __init__(self, settings: Any | None = None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._build_ui()
        self._connect_signals()
        self.set_settings(settings)

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        outer.addWidget(scroll)
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(24, 20, 24, 24)
        layout.setSpacing(12)
        scroll.setWidget(body)

        title = QLabel("Cài đặt")
        title.setObjectName("pageTitle")
        subtitle = QLabel(
            "Chọn file BAT, thư mục Output và các workbook dùng để xử lý khoản chi."
        )
        subtitle.setProperty("muted", True)
        subtitle.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(subtitle)

        card = QFrame()
        card.setProperty("card", True)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(18, 18, 18, 18)
        form = QFormLayout()
        form.setHorizontalSpacing(15)
        form.setVerticalSpacing(12)

        bat_row = QHBoxLayout()
        self.bat_edit = QLineEdit()
        self.bat_edit.setObjectName("assistantBatEdit")
        self.bat_edit.setPlaceholderText("Chọn Mo_Tro_Ly_RPA.bat")
        self.bat_edit.setClearButtonEnabled(True)
        self.browse_bat_button = QPushButton("Chọn…")
        bat_row.addWidget(self.bat_edit, 1)
        bat_row.addWidget(self.browse_bat_button)
        bat_widget = QWidget()
        bat_widget.setLayout(bat_row)
        form.addRow("File .bat mở Trợ lý ảo:", bat_widget)

        bang_ke_bat_row = QHBoxLayout()
        self.bang_ke_bat_edit = QLineEdit()
        self.bang_ke_bat_edit.setObjectName("bangKeAssistantBatEdit")
        self.bang_ke_bat_edit.setPlaceholderText("Chọn Mo_Tool_Bang_Ke.bat")
        self.bang_ke_bat_edit.setClearButtonEnabled(True)
        self.browse_bang_ke_bat_button = QPushButton("Chọn…")
        bang_ke_bat_row.addWidget(self.bang_ke_bat_edit, 1)
        bang_ke_bat_row.addWidget(self.browse_bang_ke_bat_button)
        bang_ke_bat_widget = QWidget()
        bang_ke_bat_widget.setLayout(bang_ke_bat_row)
        form.addRow("File .bat mở tool bảng kê:", bang_ke_bat_widget)

        output_row = QHBoxLayout()
        self.output_edit = QLineEdit()
        self.output_edit.setObjectName("outputDirEdit")
        self.output_edit.setClearButtonEnabled(True)
        self.browse_output_button = QPushButton("Chọn…")
        output_row.addWidget(self.output_edit, 1)
        output_row.addWidget(self.browse_output_button)
        output_widget = QWidget()
        output_widget.setLayout(output_row)
        form.addRow("Thư mục Output:", output_widget)

        rpa_expense_bat_row = QHBoxLayout()
        self.rpa_expense_bat_edit = QLineEdit()
        self.rpa_expense_bat_edit.setObjectName("rpaExpenseBatEdit")
        self.rpa_expense_bat_edit.setPlaceholderText(
            "Chọn BAT chạy PAD nhập khoản chi theo quyết toán"
        )
        self.rpa_expense_bat_edit.setClearButtonEnabled(True)
        self.browse_rpa_expense_bat_button = QPushButton("Chọn…")
        self.browse_rpa_expense_bat_button.setObjectName(
            "browseRpaExpenseBatButton"
        )
        rpa_expense_bat_row.addWidget(self.rpa_expense_bat_edit, 1)
        rpa_expense_bat_row.addWidget(self.browse_rpa_expense_bat_button)
        rpa_expense_bat_widget = QWidget()
        rpa_expense_bat_widget.setLayout(rpa_expense_bat_row)
        form.addRow("BAT RPA nhập quyết toán:", rpa_expense_bat_widget)

        daily_row = QHBoxLayout()
        self.daily_workbook_edit = QLineEdit()
        self.daily_workbook_edit.setObjectName("dailyWorkbookEdit")
        self.daily_workbook_edit.setPlaceholderText(
            "Chọn file Hàng ngày (.xlsx hoặc .xlsm)"
        )
        self.daily_workbook_edit.setClearButtonEnabled(True)
        self.browse_daily_workbook_button = QPushButton("Chọn…")
        self.browse_daily_workbook_button.setObjectName("browseDailyWorkbookButton")
        daily_row.addWidget(self.daily_workbook_edit, 1)
        daily_row.addWidget(self.browse_daily_workbook_button)
        daily_widget = QWidget()
        daily_widget.setLayout(daily_row)
        form.addRow("File Hàng ngày:", daily_widget)

        bk_row = QHBoxLayout()
        self.bk_workbook_edit = QLineEdit()
        self.bk_workbook_edit.setObjectName("bkWorkbookEdit")
        self.bk_workbook_edit.setPlaceholderText(
            "Chọn file BK Tổng hợp (.xlsx hoặc .xlsm)"
        )
        self.bk_workbook_edit.setClearButtonEnabled(True)
        self.browse_bk_workbook_button = QPushButton("Chọn…")
        self.browse_bk_workbook_button.setObjectName("browseBkWorkbookButton")
        bk_row.addWidget(self.bk_workbook_edit, 1)
        bk_row.addWidget(self.browse_bk_workbook_button)
        bk_widget = QWidget()
        bk_widget.setLayout(bk_row)
        form.addRow("File BK Tổng hợp:", bk_widget)

        payment_row = QHBoxLayout()
        self.payment_workbook_edit = QLineEdit()
        self.payment_workbook_edit.setObjectName("paymentWorkbookEdit")
        self.payment_workbook_edit.setPlaceholderText(
            "Chọn file Thanh toán Nâng hạ/VS D/O (.xlsx hoặc .xlsm)"
        )
        self.payment_workbook_edit.setClearButtonEnabled(True)
        self.browse_payment_workbook_button = QPushButton("Chọn…")
        self.browse_payment_workbook_button.setObjectName(
            "browsePaymentWorkbookButton"
        )
        payment_row.addWidget(self.payment_workbook_edit, 1)
        payment_row.addWidget(self.browse_payment_workbook_button)
        payment_widget = QWidget()
        payment_widget.setLayout(payment_row)
        form.addRow("File Thanh toán Nâng hạ/VS D/O:", payment_widget)

        # Các alias giúp phần tích hợp không phụ thuộc hậu tố ``_workbook``.
        self.daily_edit = self.daily_workbook_edit
        self.bk_edit = self.bk_workbook_edit
        self.browse_daily_button = self.browse_daily_workbook_button
        self.browse_bk_button = self.browse_bk_workbook_button
        self.payment_edit = self.payment_workbook_edit
        self.browse_payment_button = self.browse_payment_workbook_button
        for browse_button in (
            self.browse_bat_button,
            self.browse_bang_ke_bat_button,
            self.browse_output_button,
            self.browse_rpa_expense_bat_button,
            self.browse_daily_workbook_button,
            self.browse_bk_workbook_button,
            self.browse_payment_workbook_button,
        ):
            browse_button.setMinimumWidth(82)
        card_layout.addLayout(form)
        layout.addWidget(card)

        self.validation_label = QLabel()
        self.validation_label.setObjectName("settingsValidationLabel")
        self.validation_label.setWordWrap(True)
        layout.addWidget(self.validation_label)

        buttons = QHBoxLayout()
        buttons.setSpacing(8)
        self.check_button = QPushButton("Kiểm tra cấu hình")
        self.open_output_button = QPushButton("Mở thư mục Output")
        self.save_button = QPushButton("Lưu cấu hình")
        self.save_button.setObjectName("saveSettingsButton")
        self.save_button.setProperty("primary", True)
        self.check_button.setMinimumWidth(150)
        self.open_output_button.setMinimumWidth(150)
        self.save_button.setMinimumWidth(130)
        buttons.addStretch()
        buttons.addWidget(self.check_button)
        buttons.addWidget(self.open_output_button)
        buttons.addWidget(self.save_button)
        layout.addLayout(buttons)
        self.check_loading_bar = LinearLoadingBar()
        self.check_loading_bar.setAccessibleName("Tiến trình kiểm tra cấu hình")
        layout.addWidget(self.check_loading_bar)
        layout.addStretch()

    def _connect_signals(self) -> None:
        self.browse_bat_button.clicked.connect(self._browse_bat)
        self.browse_bang_ke_bat_button.clicked.connect(self._browse_bang_ke_bat)
        self.browse_output_button.clicked.connect(self._browse_output)
        self.browse_rpa_expense_bat_button.clicked.connect(
            self._browse_rpa_expense_bat
        )
        self.browse_daily_workbook_button.clicked.connect(
            self._browse_daily_workbook
        )
        self.browse_bk_workbook_button.clicked.connect(self._browse_bk_workbook)
        self.browse_payment_workbook_button.clicked.connect(
            self._browse_payment_workbook
        )
        self.save_button.clicked.connect(self._request_save)
        self.check_button.clicked.connect(
            lambda: self.check_requested.emit(self.settings_data())
        )
        self.open_output_button.clicked.connect(
            lambda: self.open_output_requested.emit(
                Path(self.output_edit.text().strip())
            )
        )
        self.bat_edit.textChanged.connect(self._validate_form)
        self.bang_ke_bat_edit.textChanged.connect(self._validate_form)
        self.output_edit.textChanged.connect(self._validate_form)
        self.rpa_expense_bat_edit.textChanged.connect(self._validate_form)
        self.daily_workbook_edit.textChanged.connect(self._validate_form)
        self.bk_workbook_edit.textChanged.connect(self._validate_form)
        self.payment_workbook_edit.textChanged.connect(self._validate_form)

    def set_settings(self, settings: Any | None) -> None:
        self.bat_edit.setText(str(_setting(settings, "assistant_bat_path") or ""))
        self.bang_ke_bat_edit.setText(
            str(_setting(settings, "bang_ke_assistant_bat_path") or "")
        )
        self.output_edit.setText(str(_setting(settings, "output_dir") or ""))
        self.daily_workbook_edit.setText(
            str(_setting(settings, "daily_workbook_path") or "")
        )
        self.bk_workbook_edit.setText(
            str(_setting(settings, "bk_workbook_path") or "")
        )
        self.payment_workbook_edit.setText(
            str(_setting(settings, "payment_workbook_path") or "")
        )
        self.rpa_expense_bat_edit.setText(
            str(_setting(settings, "rpa_expense_bat_path") or "")
        )
        self._validate_form()

    def settings_data(self) -> dict[str, Any]:
        return {
            "assistant_bat_path": self.bat_edit.text().strip(),
            "bang_ke_assistant_bat_path": self.bang_ke_bat_edit.text().strip(),
            "output_dir": self.output_edit.text().strip(),
            "daily_workbook_path": self.daily_workbook_edit.text().strip(),
            "bk_workbook_path": self.bk_workbook_edit.text().strip(),
            "payment_workbook_path": self.payment_workbook_edit.text().strip(),
            "rpa_expense_bat_path": self.rpa_expense_bat_edit.text().strip(),
        }

    values = settings_data

    def _validate_form(self, *_args: Any) -> bool:
        bat_text = self.bat_edit.text().strip()
        output_text = self.output_edit.text().strip()
        problems: list[str] = []
        if not bat_text:
            problems.append("Cần chọn file BAT mở Trợ lý ảo.")
        elif Path(bat_text).suffix.casefold() != ".bat":
            problems.append("File mở Trợ lý ảo phải có đuôi .bat.")
        elif not Path(bat_text).is_file():
            problems.append("Không tìm thấy file BAT đã chọn.")
        if not output_text:
            problems.append("Cần chọn thư mục Output.")
        bang_ke_bat = self.bang_ke_bat_edit.text().strip()
        if bang_ke_bat and Path(bang_ke_bat).suffix.casefold() != ".bat":
            problems.append("File mở tool bảng kê phải có đuôi .bat.")
        elif bang_ke_bat and not Path(bang_ke_bat).is_file():
            problems.append("Không tìm thấy file BAT mở tool bảng kê.")
        rpa_expense_bat = self.rpa_expense_bat_edit.text().strip()
        if (
            rpa_expense_bat
            and Path(rpa_expense_bat).suffix.casefold() != ".bat"
        ):
            problems.append("BAT RPA nhập quyết toán phải có đuôi .bat.")
        elif rpa_expense_bat and not Path(rpa_expense_bat).is_file():
            problems.append("Không tìm thấy BAT RPA nhập quyết toán.")
        for label, path_text in (
            ("File Hàng ngày", self.daily_workbook_edit.text().strip()),
            ("File BK Tổng hợp", self.bk_workbook_edit.text().strip()),
            (
                "File Thanh toán Nâng hạ/VS D/O",
                self.payment_workbook_edit.text().strip(),
            ),
        ):
            if path_text and Path(path_text).suffix.casefold() not in {".xlsx", ".xlsm"}:
                problems.append(f"{label} phải có đuôi .xlsx hoặc .xlsm.")

        valid = not problems
        self.save_button.setEnabled(valid)
        self.check_button.setEnabled(valid)
        self.open_output_button.setEnabled(bool(output_text))
        self.validation_label.setText(
            " • ".join(problems) if problems else "Cấu hình hợp lệ."
        )
        self.validation_label.setProperty("status", "error" if problems else "success")
        self.validation_label.style().unpolish(self.validation_label)
        self.validation_label.style().polish(self.validation_label)
        return valid

    def _request_save(self) -> None:
        if self._validate_form():
            self.save_requested.emit(self.settings_data())

    def mark_saved(self, settings: Any | None = None) -> None:
        if settings is not None:
            self.set_settings(settings)
        self.validation_label.setText("Đã lưu cấu hình.")
        self.validation_label.setProperty("status", "success")
        self.validation_label.style().unpolish(self.validation_label)
        self.validation_label.style().polish(self.validation_label)

    def show_check_result(self, success: bool, message: str) -> None:
        self.validation_label.setText(message)
        self.validation_label.setProperty("status", "success" if success else "error")
        self.validation_label.style().unpolish(self.validation_label)
        self.validation_label.style().polish(self.validation_label)

    def set_checking(self, checking: bool) -> None:
        """Hiển thị phản hồi trong lúc đọc và kiểm tra các workbook."""

        self.check_loading_bar.set_running(checking)
        set_button_loading(self.check_button, checking)
        self.check_button.setText(
            "Đang kiểm tra…" if checking else "Kiểm tra cấu hình"
        )
        self.check_button.setEnabled(
            False if checking else self.save_button.isEnabled()
        )

    def _browse_bat(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file BAT mở Trợ lý ảo",
            self.bat_edit.text().strip(),
            "Batch Windows (*.bat)",
        )
        if filename:
            self.bat_edit.setText(filename)

    def _browse_output(self) -> None:
        directory = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục Output",
            self.output_edit.text().strip(),
        )
        if directory:
            self.output_edit.setText(directory)

    def _browse_bang_ke_bat(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file BAT mở tool bảng kê",
            self.bang_ke_bat_edit.text().strip(),
            "Batch Windows (*.bat)",
        )
        if filename:
            self.bang_ke_bat_edit.setText(filename)

    def _browse_rpa_expense_bat(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn BAT chạy PAD nhập khoản chi theo quyết toán",
            self.rpa_expense_bat_edit.text().strip(),
            "Batch Windows (*.bat)",
        )
        if filename:
            self.rpa_expense_bat_edit.setText(filename)

    def _browse_daily_workbook(self) -> None:
        self._browse_workbook(
            self.daily_workbook_edit,
            "Chọn file Hàng ngày",
        )

    def _browse_bk_workbook(self) -> None:
        self._browse_workbook(
            self.bk_workbook_edit,
            "Chọn file BK Tổng hợp",
        )

    def _browse_payment_workbook(self) -> None:
        self._browse_workbook(
            self.payment_workbook_edit,
            "Chọn file Thanh toán Nâng hạ/VS D/O",
        )

    def _browse_workbook(self, target: QLineEdit, title: str) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            title,
            target.text().strip(),
            "Workbook Excel (*.xlsx *.xlsm)",
        )
        if filename:
            target.setText(filename)
