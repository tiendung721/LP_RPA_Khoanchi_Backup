"""Trang Quy trình cho luồng Trợ lý ảo -> Output -> kiểm tra JSON -> Excel."""

from __future__ import annotations

from collections.abc import Mapping
from datetime import datetime
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from .feedback import LinearLoadingBar, set_button_loading


def _get(source: Any, *names: str, default: Any = None) -> Any:
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
            return value.value if hasattr(value, "value") else value
    return default


def _saved_text(value: Any) -> str:
    if not value:
        return "—"
    if isinstance(value, datetime):
        parsed = value
    else:
        try:
            parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        except (TypeError, ValueError):
            return str(value)
    return parsed.strftime("%H:%M ngày %d/%m/%Y")


class WorkflowPage(QWidget):
    open_assistant_requested = Signal()
    open_review_requested = Signal(object)
    sync_daily_requested = Signal()
    post_expenses_requested = Signal()
    sync_payment_requested = Signal()
    run_rpa_expense_requested = Signal()

    SYNC_OPERATION = "sync"
    POSTING_OPERATION = "posting"
    PAYMENT_SYNC_OPERATION = "payment_sync"

    def __init__(
        self,
        settings: Any | None = None,
        active_batch: Any | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._settings = settings
        self._batch: Any | None = None
        self._build_ui()
        self.open_assistant_button.clicked.connect(self.open_assistant_requested)
        self.review_button.clicked.connect(
            lambda: self.open_review_requested.emit(self._batch)
        )
        self.sync_daily_button.clicked.connect(self.sync_daily_requested)
        self.post_expenses_button.clicked.connect(self.post_expenses_requested)
        self.sync_payment_button.clicked.connect(self.sync_payment_requested)
        self.run_rpa_expense_button.clicked.connect(
            self.run_rpa_expense_requested
        )
        self.set_configuration(settings)
        self.set_active_batch(active_batch)

    def _build_ui(self) -> None:
        self.setObjectName("workflowPage")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 16, 22, 18)
        layout.setSpacing(6)

        title = QLabel("Tác vụ quyết toán")
        title.setObjectName("pageTitle")
        subtitle = QLabel(
            "Chọn đúng nhóm công việc cần thực hiện. Trạng thái gần nhất luôn được hiển thị ngay tại từng tác vụ."
        )
        subtitle.setProperty("muted", True)
        subtitle.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(subtitle)

        workflow_grid = QGridLayout()
        workflow_grid.setContentsMargins(0, 8, 0, 0)
        workflow_grid.setHorizontalSpacing(12)
        workflow_grid.setVerticalSpacing(12)
        workflow_grid.setColumnStretch(0, 1)
        workflow_grid.setRowStretch(0, 1)
        workflow_grid.setRowStretch(1, 1)
        workflow_grid.setRowStretch(2, 2)
        workflow_grid.setRowStretch(3, 1)
        layout.addLayout(workflow_grid, 1)

        primary_button_width = 205

        self.step1_card = QFrame()
        self.step1_card.setProperty("card", True)
        self.step1_card.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding,
        )
        step1 = QHBoxLayout(self.step1_card)
        step1.setContentsMargins(14, 11, 14, 12)
        step1.setSpacing(16)
        step1_details = QVBoxLayout()
        step1_details.setSpacing(4)
        category1 = QLabel("BÓC TÁCH CHỨNG TỪ")
        category1.setProperty("cardCategory", True)
        step1_details.addWidget(category1)
        self.assistant_status = QLabel()
        name1 = QLabel("Bóc tách chứng từ với Trợ lý ảo")
        name1.setProperty("sectionTitle", True)
        name1.setWordWrap(True)
        step1_details.addWidget(name1)
        guide = QLabel(
            "Mở Trợ lý ảo, gửi chứng từ cần xử lý và tải file kết quả. "
            "Ứng dụng sẽ tự nhận file mới trong thư mục Output."
        )
        guide.setWordWrap(True)
        guide.setProperty("muted", True)
        step1_details.addWidget(guide)
        step1.addLayout(step1_details, 1)

        step1_controls = QVBoxLayout()
        step1_controls.setSpacing(5)
        step1_controls.addWidget(
            self.assistant_status,
            0,
            Qt.AlignmentFlag.AlignRight,
        )
        step1_controls.addStretch(1)
        self.open_assistant_button = QPushButton("Mở Trợ lý ảo")
        self.open_assistant_button.setObjectName("openAssistantButton")
        self.open_assistant_button.setProperty("primary", True)
        self.open_assistant_button.setFixedWidth(primary_button_width)
        step1_controls.addWidget(
            self.open_assistant_button,
            0,
            Qt.AlignmentFlag.AlignRight,
        )
        step1.addLayout(step1_controls)
        workflow_grid.addWidget(self.step1_card, 0, 0)

        self.step2_card = QFrame()
        self.step2_card.setProperty("card", True)
        self.step2_card.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding,
        )
        step2 = QHBoxLayout(self.step2_card)
        step2.setContentsMargins(14, 11, 14, 12)
        step2.setSpacing(16)
        step2_details = QVBoxLayout()
        step2_details.setSpacing(3)
        category2 = QLabel("KIỂM TRA DỮ LIỆU")
        category2.setProperty("cardCategory", True)
        step2_details.addWidget(category2)
        self.file_status_badge = QLabel()
        name2 = QLabel("Kiểm tra dữ liệu đã bóc tách")
        name2.setProperty("sectionTitle", True)
        name2.setWordWrap(True)
        step2_details.addWidget(name2)

        self.file_name_label = QLabel("Chưa có file bóc tách")
        self.file_name_label.setObjectName("fileNameLabel")
        self.file_name_label.setStyleSheet("font-size: 12pt; font-weight: 600;")
        self.file_name_label.setWordWrap(True)
        self.file_note_label = QLabel("Chưa nhận được file bóc tách dữ liệu.")
        self.file_note_label.setWordWrap(True)
        self.file_note_label.setProperty("muted", True)
        self.saved_label = QLabel("Lưu thành công lần cuối: —")
        self.saved_label.setProperty("muted", True)
        step2_details.addWidget(self.file_name_label)
        step2_details.addWidget(self.file_note_label)
        step2_details.addWidget(self.saved_label)
        step2.addLayout(step2_details, 1)

        step2_controls = QVBoxLayout()
        step2_controls.setSpacing(5)
        step2_controls.addWidget(
            self.file_status_badge,
            0,
            Qt.AlignmentFlag.AlignRight,
        )
        step2_controls.addStretch(1)
        self.review_button = QPushButton("Xem file bóc tách")
        self.review_button.setObjectName("openReviewButton")
        self.review_button.setProperty("primary", True)
        self.review_button.setFixedWidth(primary_button_width)
        step2_controls.addWidget(
            self.review_button,
            0,
            Qt.AlignmentFlag.AlignRight,
        )
        step2.addLayout(step2_controls)
        workflow_grid.addWidget(self.step2_card, 1, 0)

        self.step3_card = QFrame()
        self.step3_card.setObjectName("excelWorkflowCard")
        self.step3_card.setProperty("card", True)
        self.step3_card.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding,
        )
        step3 = QHBoxLayout(self.step3_card)
        step3.setContentsMargins(14, 11, 14, 12)
        step3.setSpacing(10)

        step3_intro = QWidget()
        step3_intro.setMinimumWidth(190)
        step3_intro.setMaximumWidth(220)
        step3_intro.setSizePolicy(
            QSizePolicy.Policy.Preferred,
            QSizePolicy.Policy.Expanding,
        )
        intro3 = QVBoxLayout(step3_intro)
        intro3.setContentsMargins(0, 0, 0, 0)
        intro3.setSpacing(4)
        category3 = QLabel("XỬ LÝ EXCEL")
        category3.setProperty("cardCategory", True)
        intro3.addWidget(category3)
        name3 = QLabel("Xử lý và đồng bộ dữ liệu Excel")
        name3.setProperty("sectionTitle", True)
        name3.setWordWrap(True)
        intro3.addWidget(name3)

        self.step3_card.setToolTip(
            "Mỗi tác vụ sử dụng một luồng dữ liệu riêng và hiển thị "
            "kết quả gần nhất ngay trong cùng hàng."
        )

        self.sync_status_label = QLabel("Đồng bộ gần nhất: —")
        self.sync_status_label.setObjectName("dailySyncStatusLabel")
        self.sync_status_label.setWordWrap(True)
        self.sync_status_label.setProperty("muted", True)
        self.sync_status_label.setMinimumWidth(0)
        self.sync_status_label.setSizePolicy(
            QSizePolicy.Policy.Ignored,
            QSizePolicy.Policy.Preferred,
        )
        self.posting_status_label = QLabel("Nhập khoản chi gần nhất: —")
        self.posting_status_label.setObjectName("expensePostingStatusLabel")
        self.posting_status_label.setWordWrap(True)
        self.posting_status_label.setProperty("muted", True)
        self.posting_status_label.setMinimumWidth(0)
        self.posting_status_label.setSizePolicy(
            QSizePolicy.Policy.Ignored,
            QSizePolicy.Policy.Preferred,
        )
        self.payment_sync_status_label = QLabel(
            "Đồng bộ BK → Thanh toán gần nhất: —"
        )
        self.payment_sync_status_label.setObjectName("paymentSyncStatusLabel")
        self.payment_sync_status_label.setWordWrap(True)
        self.payment_sync_status_label.setProperty("muted", True)
        self.payment_sync_status_label.setMinimumWidth(0)
        self.payment_sync_status_label.setSizePolicy(
            QSizePolicy.Policy.Ignored,
            QSizePolicy.Policy.Preferred,
        )
        self.sync_daily_button = QPushButton("Đồng bộ dữ liệu Hàng ngày")
        self.sync_daily_button.setObjectName("syncDailyWorkbookButton")
        self.sync_daily_button.setProperty("primary", True)
        self.post_expenses_button = QPushButton("Nhập khoản chi vào BK")
        self.post_expenses_button.setObjectName("postExpensesWorkbookButton")
        self.post_expenses_button.setProperty("primary", True)
        self.sync_payment_button = QPushButton("Đồng bộ BK → Thanh toán")
        self.sync_payment_button.setObjectName("syncPaymentWorkbookButton")
        self.sync_payment_button.setProperty("primary", True)
        for button in (
            self.sync_daily_button,
            self.post_expenses_button,
            self.sync_payment_button,
        ):
            button.setMinimumWidth(0)
            button.setMaximumWidth(205)
            button.setSizePolicy(
                QSizePolicy.Policy.Ignored,
                QSizePolicy.Policy.Fixed,
        )
        self.excel_loading_bar = LinearLoadingBar()
        self.excel_loading_bar.setAccessibleName("Tiến trình xử lý Excel")
        intro3.addWidget(self.excel_loading_bar)
        intro3.addStretch(1)
        step3.addWidget(step3_intro)

        def add_excel_action(
            title_text: str,
            description: str,
            status_label: QLabel,
            button: QPushButton,
        ) -> QFrame:
            action = QFrame()
            action.setProperty("actionRow", True)
            action.setToolTip(description)
            action_layout = QVBoxLayout(action)
            action_layout.setContentsMargins(10, 6, 10, 6)
            action_layout.setSpacing(5)
            details = QVBoxLayout()
            details.setSpacing(1)
            action_title = QLabel(title_text)
            action_title.setProperty("actionTitle", True)
            action_title.setToolTip(description)
            action_title.setMinimumWidth(0)
            action_title.setSizePolicy(
                QSizePolicy.Policy.Ignored,
                QSizePolicy.Policy.Preferred,
            )
            details.addWidget(action_title)
            details.addWidget(status_label)
            action_layout.addLayout(details)
            action_layout.addWidget(
                button,
                0,
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
            )
            step3.addWidget(action, 1)
            return action

        self.daily_sync_action = add_excel_action(
            "Hàng ngày → BK",
            "Cập nhật dữ liệu từ workbook Hàng ngày vào BK Tổng hợp.",
            self.sync_status_label,
            self.sync_daily_button,
        )
        self.expense_posting_action = add_excel_action(
            "Khoản chi → BK",
            "Ghi các khoản chi đã kiểm tra và xác nhận vào BK Tổng hợp.",
            self.posting_status_label,
            self.post_expenses_button,
        )
        self.payment_sync_action = add_excel_action(
            "BK → Thanh toán",
            "Chuyển dữ liệu từ BK Tổng hợp sang workbook Thanh toán.",
            self.payment_sync_status_label,
            self.sync_payment_button,
        )
        workflow_grid.addWidget(self.step3_card, 2, 0)

        self.step4_card = QFrame()
        self.step4_card.setObjectName("rpaExpenseWorkflowCard")
        self.step4_card.setProperty("card", True)
        self.step4_card.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding,
        )
        step4 = QHBoxLayout(self.step4_card)
        step4.setContentsMargins(14, 11, 14, 12)
        step4.setSpacing(16)
        step4_details = QVBoxLayout()
        step4_details.setSpacing(4)
        category4 = QLabel("TỰ ĐỘNG HÓA RPA")
        category4.setProperty("cardCategory", True)
        step4_details.addWidget(category4)
        name4 = QLabel("Nhập khoản chi lên phần mềm quyết toán")
        name4.setProperty("sectionTitle", True)
        name4.setWordWrap(True)
        step4_details.addWidget(name4)
        guide4 = QLabel(
            "Chọn sheet BK, chọn một hoặc nhiều số quyết toán, sau đó khởi chạy "
            "flow PAD để nhập các khoản chi đã tổng hợp lên phần mềm quyết toán."
        )
        guide4.setWordWrap(True)
        guide4.setProperty("muted", True)
        step4_details.addWidget(guide4)
        step4.addLayout(step4_details, 1)

        step4_controls = QVBoxLayout()
        step4_controls.setSpacing(5)
        self.rpa_configuration_status = QLabel()
        step4_controls.addWidget(
            self.rpa_configuration_status,
            0,
            Qt.AlignmentFlag.AlignRight,
        )
        self.run_rpa_expense_button = QPushButton("Nhập PM quyết toán")
        self.run_rpa_expense_button.setObjectName("runRpaExpenseButton")
        self.run_rpa_expense_button.setProperty("primary", True)
        self.run_rpa_expense_button.setFixedWidth(primary_button_width)
        step4_controls.addWidget(self.run_rpa_expense_button)
        self.rpa_loading_bar = LinearLoadingBar()
        self.rpa_loading_bar.setAccessibleName("Tiến trình chuẩn bị RPA")
        step4_controls.addWidget(self.rpa_loading_bar)
        self.rpa_expense_status_label = QLabel("Chạy RPA gần nhất: —")
        self.rpa_expense_status_label.setObjectName("rpaExpenseStatusLabel")
        self.rpa_expense_status_label.setWordWrap(True)
        self.rpa_expense_status_label.setProperty("muted", True)
        step4_controls.addWidget(self.rpa_expense_status_label)
        step4.addLayout(step4_controls)
        workflow_grid.addWidget(self.step4_card, 3, 0)

        # Aliases có chủ ý để lớp điều phối có thể dùng cách đặt tên theo nghiệp vụ
        # mà không tạo thêm widget/nút chính.
        self.daily_sync_button = self.sync_daily_button
        self.expense_posting_button = self.post_expenses_button
        self.daily_sync_status_label = self.sync_status_label
        self.expense_posting_status_label = self.posting_status_label
        self._excel_running_operation: str | None = None
        self._rpa_running = False
        self._rpa_button_text = self.run_rpa_expense_button.text()
        self._excel_button_texts = {
            self.SYNC_OPERATION: self.sync_daily_button.text(),
            self.POSTING_OPERATION: self.post_expenses_button.text(),
            self.PAYMENT_SYNC_OPERATION: self.sync_payment_button.text(),
        }

    def set_configuration(self, settings: Any | None) -> None:
        self._settings = settings
        bat = str(_get(settings, "assistant_bat_path", default="") or "").strip()
        if bat:
            self.assistant_status.setText("Đã cấu hình")
            self.assistant_status.setStyleSheet(
                "color: #15803D; background: #ECFDF3; border-radius: 5px; "
                "padding: 4px 8px;"
            )
        else:
            self.assistant_status.setText("Chưa cấu hình BAT")
            self.assistant_status.setStyleSheet(
                "color: #A16207; background: #FFF8DB; border-radius: 5px; "
                "padding: 4px 8px;"
            )
        rpa_bat = str(
            _get(settings, "rpa_expense_bat_path", default="") or ""
        ).strip()
        if rpa_bat:
            self.rpa_configuration_status.setText("Đã cấu hình BAT RPA")
            self.rpa_configuration_status.setStyleSheet(
                "color: #15803D; background: #ECFDF3; border-radius: 5px; "
                "padding: 4px 8px;"
            )
        else:
            self.rpa_configuration_status.setText("Chưa cấu hình BAT RPA")
            self.rpa_configuration_status.setStyleSheet(
                "color: #A16207; background: #FFF8DB; border-radius: 5px; "
                "padding: 4px 8px;"
            )

    def set_active_batch(self, batch: Any | None) -> None:
        self._batch = batch
        metadata = _get(batch, "metadata", default=batch)
        has_batch = metadata is not None and _get(metadata, "id", "batch_id") is not None
        if not has_batch:
            self.file_status_badge.setText("Chưa có file")
            self.file_status_badge.setStyleSheet(
                "color: #64748B; background: #E2E8F0; border-radius: 5px; "
                "padding: 4px 8px;"
            )
            self.file_name_label.setText("Chưa có file bóc tách")
            self.file_note_label.setText("Chưa nhận được file bóc tách dữ liệu.")
            self.saved_label.setText("Lưu thành công lần cuối: —")
            self.review_button.setEnabled(False)
            return

        filename = str(
            _get(metadata, "source_filename", "filename", default="ket_qua_boc_tach.json")
        )
        status = str(_get(metadata, "status", default="RECEIVED")).split(".")[-1].upper()
        invalid = status == "INVALID"
        self.file_name_label.setText(filename)
        self.file_name_label.setToolTip(filename)
        if invalid:
            self.file_status_badge.setText("File không hợp lệ")
            self.file_status_badge.setStyleSheet(
                "color: #B42318; background: #FFF0F0; border-radius: 5px; "
                "padding: 4px 8px; font-weight: 600;"
            )
            self.file_note_label.setText(
                str(_get(metadata, "last_error", default="File JSON không hợp lệ."))
            )
            self.saved_label.setText("Lưu thành công lần cuối: —")
            self.review_button.setEnabled(False)
            return

        self.file_status_badge.setText("Đã có file")
        self.file_status_badge.setStyleSheet(
            "color: #15803D; background: #ECFDF3; border-radius: 5px; "
            "padding: 4px 8px; font-weight: 600;"
        )
        saved_at = _get(metadata, "last_saved_at", "saved_at")
        self.file_note_label.setText(
            "Đã lưu dữ liệu bóc tách JSON sau khi kiểm tra."
            if saved_at
            else "Đã nhận file mới; hãy kiểm tra và lưu dữ liệu."
        )
        self.saved_label.setText(
            f"Lưu thành công lần cuối: {_saved_text(saved_at)}"
        )
        self.review_button.setEnabled(True)

    def clear_active_batch(self) -> None:
        self.set_active_batch(None)

    @staticmethod
    def _excel_operation(operation: Any) -> str:
        value = str(getattr(operation, "value", operation) or "").casefold()
        if value in {"sync", "daily_sync", "sync_daily", "daily"}:
            return WorkflowPage.SYNC_OPERATION
        if value in {
            "posting",
            "post",
            "expense_posting",
            "post_expenses",
            "expenses",
        }:
            return WorkflowPage.POSTING_OPERATION
        if value in {
            "payment_sync",
            "sync_payment",
            "bk_to_payment",
            "payment",
        }:
            return WorkflowPage.PAYMENT_SYNC_OPERATION
        raise ValueError(f"Nghiệp vụ Excel không hợp lệ: {operation!r}")

    def set_excel_running(self, operation: Any, message: str = "") -> None:
        """Hiển thị tiến độ và khóa đồng thời cả hai thao tác Excel."""

        normalized = self._excel_operation(operation)
        self._excel_running_operation = normalized
        self.excel_loading_bar.set_running(True)
        for button in (
            self.sync_daily_button,
            self.post_expenses_button,
            self.sync_payment_button,
        ):
            set_button_loading(button, False)
        self.sync_daily_button.setEnabled(False)
        self.post_expenses_button.setEnabled(False)
        self.sync_payment_button.setEnabled(False)
        self.run_rpa_expense_button.setEnabled(False)
        if normalized == self.SYNC_OPERATION:
            set_button_loading(self.sync_daily_button, True)
            self.sync_daily_button.setText("Đang đồng bộ…")
            self.sync_status_label.setText(
                f"Đồng bộ: {message or 'Đang phân tích dữ liệu…'}"
            )
        elif normalized == self.POSTING_OPERATION:
            set_button_loading(self.post_expenses_button, True)
            self.post_expenses_button.setText("Đang nhập…")
            self.posting_status_label.setText(
                f"Nhập khoản chi: {message or 'Đang phân tích dữ liệu…'}"
            )
        else:
            set_button_loading(self.sync_payment_button, True)
            self.sync_payment_button.setText("Đang đồng bộ…")
            self.payment_sync_status_label.setText(
                "Đồng bộ BK → Thanh toán: "
                f"{message or 'Đang phân tích dữ liệu…'}"
            )

    def set_excel_progress(self, operation: Any, message: str) -> None:
        normalized = self._excel_operation(operation)
        if normalized == self.SYNC_OPERATION:
            self.sync_status_label.setText(f"Đồng bộ: {message}")
        elif normalized == self.POSTING_OPERATION:
            self.posting_status_label.setText(f"Nhập khoản chi: {message}")
        else:
            self.payment_sync_status_label.setText(
                f"Đồng bộ BK → Thanh toán: {message}"
            )

    def set_excel_result(self, operation: Any, result: Any = None) -> None:
        """Cập nhật kết quả gần nhất từ dataclass, mapping hoặc chuỗi."""

        normalized = self._excel_operation(operation)
        if isinstance(result, str):
            message = result
        else:
            message = str(
                _get(result, "message", "summary", default="Hoàn tất.") or "Hoàn tất."
            )
        if normalized == self.SYNC_OPERATION:
            self.sync_status_label.setText(f"Đồng bộ gần nhất: {message}")
        elif normalized == self.POSTING_OPERATION:
            self.posting_status_label.setText(
                f"Nhập khoản chi gần nhất: {message}"
            )
        else:
            self.payment_sync_status_label.setText(
                f"Đồng bộ BK → Thanh toán gần nhất: {message}"
            )

    def set_excel_idle(self, operation: Any | None = None) -> None:
        """Khôi phục hai nút sau khi controller phát ``finished``."""

        if operation is not None:
            self._excel_operation(operation)
        self._excel_running_operation = None
        self.excel_loading_bar.set_running(False)
        for button in (
            self.sync_daily_button,
            self.post_expenses_button,
            self.sync_payment_button,
        ):
            set_button_loading(button, False)
        self.sync_daily_button.setText(
            self._excel_button_texts[self.SYNC_OPERATION]
        )
        self.post_expenses_button.setText(
            self._excel_button_texts[self.POSTING_OPERATION]
        )
        self.sync_payment_button.setText(
            self._excel_button_texts[self.PAYMENT_SYNC_OPERATION]
        )
        enabled = not self._rpa_running
        self.sync_daily_button.setEnabled(enabled)
        self.post_expenses_button.setEnabled(enabled)
        self.sync_payment_button.setEnabled(enabled)
        self.run_rpa_expense_button.setEnabled(enabled)

    def set_excel_actions_enabled(self, enabled: bool) -> None:
        if self._excel_running_operation is not None and enabled:
            return
        self.sync_daily_button.setEnabled(enabled)
        self.post_expenses_button.setEnabled(enabled)
        self.sync_payment_button.setEnabled(enabled)
        self.run_rpa_expense_button.setEnabled(enabled and not self._rpa_running)

    def set_rpa_running(self, message: str = "") -> None:
        self._rpa_running = True
        self.rpa_loading_bar.set_running(True)
        set_button_loading(self.run_rpa_expense_button, True)
        self.sync_daily_button.setEnabled(False)
        self.post_expenses_button.setEnabled(False)
        self.sync_payment_button.setEnabled(False)
        self.run_rpa_expense_button.setEnabled(False)
        self.run_rpa_expense_button.setText("Đang chuẩn bị RPA…")
        self.rpa_expense_status_label.setText(
            f"RPA: {message or 'Đang chuẩn bị dữ liệu…'}"
        )

    def set_rpa_progress(self, message: str) -> None:
        self.rpa_expense_status_label.setText(f"RPA: {message}")

    def set_rpa_result(self, result: Any = None) -> None:
        message = (
            result
            if isinstance(result, str)
            else _get(result, "message", default="Đã khởi chạy PAD.")
        )
        self.rpa_expense_status_label.setText(
            f"Chạy RPA gần nhất: {message or 'Đã khởi chạy PAD.'}"
        )

    def set_rpa_idle(self) -> None:
        self._rpa_running = False
        self.rpa_loading_bar.set_running(False)
        set_button_loading(self.run_rpa_expense_button, False)
        self.run_rpa_expense_button.setText(self._rpa_button_text)
        if self._excel_running_operation is None:
            self.sync_daily_button.setEnabled(True)
            self.post_expenses_button.setEnabled(True)
            self.sync_payment_button.setEnabled(True)
            self.run_rpa_expense_button.setEnabled(True)

    @property
    def rpa_running(self) -> bool:
        return self._rpa_running

    @property
    def excel_running_operation(self) -> str | None:
        return self._excel_running_operation

    @property
    def active_batch(self) -> Any | None:
        return self._batch
