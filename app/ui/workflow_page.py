"""Trang Quy trình cho luồng Trợ lý ảo -> Output -> kiểm tra JSON -> Excel."""

from __future__ import annotations

from collections.abc import Mapping
from datetime import datetime
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QBoxLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
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
    open_bang_ke_assistant_requested = Signal()
    open_review_requested = Signal(object)
    sync_daily_requested = Signal()
    post_expenses_requested = Signal()
    sync_payment_requested = Signal()
    export_carrier_requested = Signal()
    export_posting_requested = Signal()
    run_rpa_expense_requested = Signal()
    view_latest_excel_requested = Signal(str)
    view_latest_rpa_requested = Signal()

    SYNC_OPERATION = "sync"
    POSTING_OPERATION = "posting"
    PAYMENT_SYNC_OPERATION = "payment_sync"
    CARRIER_EXPORT_OPERATION = "carrier_export"
    POSTING_EXPORT_OPERATION = "posting_export"

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
        self.open_bang_ke_assistant_button.clicked.connect(
            self.open_bang_ke_assistant_requested
        )
        self.review_button.clicked.connect(
            lambda: self.open_review_requested.emit(self._batch)
        )
        self.sync_daily_button.clicked.connect(self.sync_daily_requested)
        self.post_expenses_button.clicked.connect(self.post_expenses_requested)
        self.sync_payment_button.clicked.connect(self.sync_payment_requested)
        self.export_carrier_button.clicked.connect(self.export_carrier_requested)
        self.export_posting_button.clicked.connect(self.export_posting_requested)
        self.view_posting_export_button.clicked.connect(
            lambda: self.view_latest_excel_requested.emit(self.POSTING_EXPORT_OPERATION)
        )
        self.run_rpa_expense_button.clicked.connect(
            self.run_rpa_expense_requested
        )
        self.view_daily_sync_button.clicked.connect(
            lambda: self.view_latest_excel_requested.emit(self.SYNC_OPERATION)
        )
        self.view_expense_posting_button.clicked.connect(
            lambda: self.view_latest_excel_requested.emit(self.POSTING_OPERATION)
        )
        self.view_payment_sync_button.clicked.connect(
            lambda: self.view_latest_excel_requested.emit(self.PAYMENT_SYNC_OPERATION)
        )
        self.view_carrier_export_button.clicked.connect(
            lambda: self.view_latest_excel_requested.emit(self.CARRIER_EXPORT_OPERATION)
        )
        self.view_rpa_expense_button.clicked.connect(
            self.view_latest_rpa_requested
        )
        self.set_configuration(settings)
        self.set_active_batch(active_batch)

    def _build_ui(self) -> None:
        self.setObjectName("workflowPage")
        page_layout = QVBoxLayout(self)
        page_layout.setContentsMargins(0, 0, 0, 0)
        self.workflow_scroll = QScrollArea()
        self.workflow_scroll.setObjectName("workflowScrollArea")
        self.workflow_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.workflow_scroll.setWidgetResizable(True)
        self.workflow_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        content = QWidget()
        content.setObjectName("workflowContent")
        layout = QVBoxLayout(content)
        layout.setContentsMargins(22, 16, 22, 18)
        layout.setSpacing(6)
        self.workflow_scroll.setWidget(content)
        page_layout.addWidget(self.workflow_scroll)

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
            "Ứng dụng sẽ tự nhận file kết quả khi bạn tải về."
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
        self.open_bang_ke_assistant_button = QPushButton("Mở trợ lý bảng kê")
        self.open_bang_ke_assistant_button.setObjectName("openBangKeAssistantButton")
        self.open_bang_ke_assistant_button.setFixedWidth(primary_button_width)
        assistant_buttons = QHBoxLayout()
        assistant_buttons.setSpacing(8)
        assistant_buttons.addWidget(self.open_assistant_button)
        assistant_buttons.addWidget(self.open_bang_ke_assistant_button)
        step1_controls.addLayout(assistant_buttons)
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
        step3 = QVBoxLayout(self.step3_card)
        step3.setContentsMargins(14, 11, 14, 12)
        step3.setSpacing(3)
        category3 = QLabel("XỬ LÝ EXCEL")
        category3.setProperty("cardCategory", True)
        step3.addWidget(category3)
        name3 = QLabel("Xử lý và đồng bộ dữ liệu Excel")
        name3.setProperty("sectionTitle", True)
        name3.setWordWrap(True)
        step3.addWidget(name3)

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
        self.carrier_export_status_label = QLabel("Xuất theo bên VT gần nhất: —")
        self.carrier_export_status_label.setObjectName("carrierExportStatusLabel")
        self.carrier_export_status_label.setWordWrap(True)
        self.carrier_export_status_label.setProperty("muted", True)
        self.carrier_export_status_label.setMinimumWidth(0)
        self.carrier_export_status_label.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred
        )
        self.posting_export_status_label = QLabel("Chưa có lượt nhập BK để xuất")
        self.posting_export_status_label.setObjectName("postingExportStatusLabel")
        self.posting_export_status_label.setWordWrap(True)
        self.posting_export_status_label.setProperty("muted", True)
        self.posting_export_status_label.setMinimumWidth(0)
        self.sync_daily_button = QPushButton("Đồng bộ")
        self.sync_daily_button.setObjectName("syncDailyWorkbookButton")
        self.sync_daily_button.setProperty("primary", True)
        self.sync_daily_button.setAccessibleName("Đồng bộ dữ liệu Hàng ngày vào BK")
        self.post_expenses_button = QPushButton("Nhập vào BK")
        self.post_expenses_button.setObjectName("postExpensesWorkbookButton")
        self.post_expenses_button.setProperty("primary", True)
        self.post_expenses_button.setAccessibleName("Nhập khoản chi vào BK")
        self.sync_payment_button = QPushButton("Đồng bộ")
        self.sync_payment_button.setObjectName("syncPaymentWorkbookButton")
        self.sync_payment_button.setProperty("primary", True)
        self.sync_payment_button.setAccessibleName("Đồng bộ BK sang Thanh toán")
        self.export_carrier_button = QPushButton("Xuất Excel")
        self.export_carrier_button.setObjectName("exportCarrierWorkbookButton")
        self.export_carrier_button.setProperty("primary", True)
        self.export_carrier_button.setAccessibleName("Xuất phí trong Nam theo bên vận tải từ BK")
        self.export_posting_button = QPushButton("Xuất Excel")
        self.export_posting_button.setObjectName("exportPostingWorkbookButton")
        self.export_posting_button.setProperty("primary", True)
        self.export_posting_button.setAccessibleName("Xuất các khoản chi của lượt nhập BK gần nhất")
        self.export_posting_button.setEnabled(False)
        self.view_daily_sync_button = QPushButton("Chưa có dữ liệu")
        self.view_daily_sync_button.setObjectName("viewLatestDailySyncButton")
        self.view_expense_posting_button = QPushButton("Chưa có dữ liệu")
        self.view_expense_posting_button.setObjectName(
            "viewLatestExpensePostingButton"
        )
        self.view_payment_sync_button = QPushButton("Chưa có dữ liệu")
        self.view_payment_sync_button.setObjectName("viewLatestPaymentSyncButton")
        self.view_carrier_export_button = QPushButton("Chưa có file")
        self.view_carrier_export_button.setObjectName("viewLatestCarrierExportButton")
        self.view_posting_export_button = QPushButton("Chưa có file")
        self.view_posting_export_button.setObjectName("viewLatestPostingExportButton")
        for view_button in (
            self.view_daily_sync_button,
            self.view_expense_posting_button,
            self.view_payment_sync_button,
            self.view_carrier_export_button,
            self.view_posting_export_button,
        ):
            view_button.setProperty("link", True)
            view_button.setEnabled(False)
            view_button.setFixedWidth(105)
            view_button.setSizePolicy(
                QSizePolicy.Policy.Fixed,
                QSizePolicy.Policy.Fixed,
            )
        for button in (
            self.sync_daily_button,
            self.post_expenses_button,
            self.sync_payment_button,
            self.export_carrier_button,
            self.export_posting_button,
        ):
            button.setFixedWidth(140)
            button.setMinimumHeight(38)
            button.setSizePolicy(
                QSizePolicy.Policy.Fixed,
                QSizePolicy.Policy.Fixed,
        )
        self.excel_loading_bar = LinearLoadingBar()
        self.excel_loading_bar.setAccessibleName("Tiến trình xử lý Excel")
        step3.addWidget(self.excel_loading_bar)
        step3.addSpacing(4)
        self._excel_action_layouts: list[tuple[QBoxLayout, QFrame]] = []

        def add_excel_action(
            title_text: str,
            description: str,
            status_label: QLabel,
            button: QPushButton,
            view_button: QPushButton,
        ) -> QFrame:
            action = QFrame()
            action.setProperty("taskRow", True)
            action.setToolTip(description)
            action_layout = QHBoxLayout(action)
            action_layout.setContentsMargins(10, 5, 8, 5)
            action_layout.setSpacing(8)
            action_title = QLabel(title_text)
            action_title.setProperty("actionTitle", True)
            action_title.setToolTip(description)
            action_title.setFixedWidth(145)
            action_title.setSizePolicy(
                QSizePolicy.Policy.Preferred,
                QSizePolicy.Policy.Preferred,
            )
            status_label.setToolTip(status_label.text())
            status_label.setSizePolicy(
                QSizePolicy.Policy.Ignored,
                QSizePolicy.Policy.Preferred,
            )
            details_layout = QBoxLayout(QBoxLayout.Direction.LeftToRight)
            details_layout.setContentsMargins(0, 0, 0, 0)
            details_layout.setSpacing(8)
            details_layout.addWidget(action_title)
            details_layout.addWidget(status_label, 1)
            action_layout.addLayout(details_layout, 1)
            action_layout.addWidget(button, 0, Qt.AlignmentFlag.AlignVCenter)
            action_layout.addWidget(view_button, 0, Qt.AlignmentFlag.AlignVCenter)
            action.setMinimumHeight(52)
            step3.addWidget(action)
            self._excel_action_layouts.append((details_layout, action))
            return action

        self.daily_sync_action = add_excel_action(
            "Hàng ngày → BK",
            "Cập nhật dữ liệu từ file Hàng ngày vào BK Tổng hợp.",
            self.sync_status_label,
            self.sync_daily_button,
            self.view_daily_sync_button,
        )
        self.expense_posting_action = add_excel_action(
            "Khoản chi → BK",
            "Ghi các khoản chi đã kiểm tra và xác nhận vào BK Tổng hợp.",
            self.posting_status_label,
            self.post_expenses_button,
            self.view_expense_posting_button,
        )
        self.posting_export_action = add_excel_action(
            "Kết quả nhập BK → Excel",
            "Xuất tất cả khoản chi của lượt nhập BK gần nhất, kèm file nguồn và trạng thái.",
            self.posting_export_status_label,
            self.export_posting_button,
            self.view_posting_export_button,
        )
        self.payment_sync_action = add_excel_action(
            "BK → Thanh toán",
            "Chuyển dữ liệu từ BK Tổng hợp sang file Thanh toán.",
            self.payment_sync_status_label,
            self.sync_payment_button,
            self.view_payment_sync_button,
        )
        self.carrier_export_action = add_excel_action(
            "BK → Bên VT",
            "Xuất các khoản phí trong Nam của một bên vận tải từ những sheet tháng BK đã chọn.",
            self.carrier_export_status_label,
            self.export_carrier_button,
            self.view_carrier_export_button,
        )
        self._excel_compact = False
        self._set_excel_compact(True)
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
        category4 = QLabel("NHẬP DỮ LIỆU QUYẾT TOÁN")
        category4.setProperty("cardCategory", True)
        step4_details.addWidget(category4)
        name4 = QLabel("Nhập khoản chi lên phần mềm quyết toán")
        name4.setProperty("sectionTitle", True)
        name4.setWordWrap(True)
        step4_details.addWidget(name4)
        guide4 = QLabel(
            "Chọn tháng BK và các số quyết toán cần nhập. Ứng dụng sẽ mở công cụ "
            "tự động để chuyển các khoản chi sang phần mềm quyết toán."
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
        self.rpa_expense_status_label = QLabel("Lần nhập gần nhất: —")
        self.rpa_expense_status_label.setObjectName("rpaExpenseStatusLabel")
        self.rpa_expense_status_label.setWordWrap(True)
        self.rpa_expense_status_label.setProperty("muted", True)
        self.view_rpa_expense_button = QPushButton("Chưa có dữ liệu đã gửi")
        self.view_rpa_expense_button.setObjectName("viewLatestRpaExpenseButton")
        self.view_rpa_expense_button.setProperty("link", True)
        self.view_rpa_expense_button.setMaximumWidth(165)
        self.view_rpa_expense_button.setEnabled(False)
        rpa_recent = QHBoxLayout()
        rpa_recent.setContentsMargins(0, 0, 0, 0)
        rpa_recent.setSpacing(8)
        rpa_recent.addWidget(self.rpa_expense_status_label, 1)
        rpa_recent.addWidget(self.view_rpa_expense_button)
        step4_details.addLayout(rpa_recent)
        self.rpa_loading_bar = LinearLoadingBar()
        self.rpa_loading_bar.setAccessibleName("Tiến trình chuẩn bị nhập dữ liệu")
        step4_controls.addWidget(self.rpa_loading_bar)
        step4.addLayout(step4_controls)
        workflow_grid.addWidget(self.step4_card, 3, 0)

        # Aliases có chủ ý để lớp điều phối có thể dùng cách đặt tên theo nghiệp vụ
        # mà không tạo thêm widget/nút chính.
        self.daily_sync_button = self.sync_daily_button
        self.expense_posting_button = self.post_expenses_button
        self.daily_sync_status_label = self.sync_status_label
        self.expense_posting_status_label = self.posting_status_label
        self._excel_running_operation: str | None = None
        self._posting_export_available = False
        self._rpa_running = False
        self._rpa_button_text = self.run_rpa_expense_button.text()
        self._excel_button_texts = {
            self.SYNC_OPERATION: self.sync_daily_button.text(),
            self.POSTING_OPERATION: self.post_expenses_button.text(),
            self.PAYMENT_SYNC_OPERATION: self.sync_payment_button.text(),
            self.CARRIER_EXPORT_OPERATION: self.export_carrier_button.text(),
            self.POSTING_EXPORT_OPERATION: self.export_posting_button.text(),
        }

    def resizeEvent(self, event: Any) -> None:
        super().resizeEvent(event)
        self._set_excel_compact(self.width() < 980)

    def _set_excel_compact(self, compact: bool) -> None:
        if self._excel_compact == compact:
            return
        self._excel_compact = compact
        for details_layout, action in self._excel_action_layouts:
            if compact:
                details_layout.setDirection(QBoxLayout.Direction.TopToBottom)
                details_layout.setSpacing(2)
                action.setMinimumHeight(72)
            else:
                details_layout.setDirection(QBoxLayout.Direction.LeftToRight)
                details_layout.setSpacing(8)
                action.setMinimumHeight(52)

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
            self.assistant_status.setText("Chưa thiết lập trợ lý")
            self.assistant_status.setStyleSheet(
                "color: #A16207; background: #FFF8DB; border-radius: 5px; "
                "padding: 4px 8px;"
            )
        bang_ke_bat = str(
            _get(settings, "bang_ke_assistant_bat_path", default="") or ""
        ).strip()
        self.open_bang_ke_assistant_button.setToolTip(
            "" if bang_ke_bat else "Chưa thiết lập trợ lý bảng kê trong Cài đặt."
        )
        rpa_bat = str(
            _get(settings, "rpa_expense_bat_path", default="") or ""
        ).strip()
        if rpa_bat:
            self.rpa_configuration_status.setText("Đã sẵn sàng")
            self.rpa_configuration_status.setStyleSheet(
                "color: #15803D; background: #ECFDF3; border-radius: 5px; "
                "padding: 4px 8px;"
            )
        else:
            self.rpa_configuration_status.setText("Cần thiết lập công cụ nhập")
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
            self.file_note_label.setToolTip("")
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
                "Không đọc được file kết quả. Hãy kiểm tra file vừa tải về và thử lại."
            )
            self.file_note_label.setToolTip(
                str(_get(metadata, "last_error", default=""))
            )
            self.saved_label.setText("Lưu thành công lần cuối: —")
            self.review_button.setEnabled(False)
            return

        self.file_status_badge.setText("Đã có file")
        self.file_note_label.setToolTip("")
        self.file_status_badge.setStyleSheet(
            "color: #15803D; background: #ECFDF3; border-radius: 5px; "
            "padding: 4px 8px; font-weight: 600;"
        )
        saved_at = _get(metadata, "last_saved_at", "saved_at")
        self.file_note_label.setText(
            "Đã kiểm tra và lưu dữ liệu bóc tách."
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
        if value in {"carrier_export", "export_carrier", "xlsx_by_carrier"}:
            return WorkflowPage.CARRIER_EXPORT_OPERATION
        if value in {"posting_export", "export_posting", "expense_export"}:
            return WorkflowPage.POSTING_EXPORT_OPERATION
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
            self.export_carrier_button,
            self.export_posting_button,
        ):
            set_button_loading(button, False)
        self.sync_daily_button.setEnabled(False)
        self.post_expenses_button.setEnabled(False)
        self.sync_payment_button.setEnabled(False)
        self.export_carrier_button.setEnabled(False)
        self.export_posting_button.setEnabled(False)
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
        elif normalized == self.PAYMENT_SYNC_OPERATION:
            set_button_loading(self.sync_payment_button, True)
            self.sync_payment_button.setText("Đang đồng bộ…")
            self.payment_sync_status_label.setText(
                "Đồng bộ BK → Thanh toán: "
                f"{message or 'Đang phân tích dữ liệu…'}"
            )
        elif normalized == self.CARRIER_EXPORT_OPERATION:
            set_button_loading(self.export_carrier_button, True)
            self.export_carrier_button.setText("Đang xuất…")
            self.carrier_export_status_label.setText(message or "Đang đọc dữ liệu BK…")
        else:
            set_button_loading(self.export_posting_button, True)
            self.export_posting_button.setText("Đang xuất…")
            self.posting_export_status_label.setText(message or "Đang dựng file Excel…")

    def set_excel_progress(self, operation: Any, message: str) -> None:
        normalized = self._excel_operation(operation)
        if normalized == self.SYNC_OPERATION:
            self.sync_status_label.setText(f"Đồng bộ: {message}")
        elif normalized == self.POSTING_OPERATION:
            self.posting_status_label.setText(f"Nhập khoản chi: {message}")
        elif normalized == self.PAYMENT_SYNC_OPERATION:
            self.payment_sync_status_label.setText(
                f"Đồng bộ BK → Thanh toán: {message}"
            )
        elif normalized == self.CARRIER_EXPORT_OPERATION:
            self.carrier_export_status_label.setText(f"Xuất theo bên VT: {message}")
        else:
            self.posting_export_status_label.setText(f"Xuất kết quả nhập BK: {message}")

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
        elif normalized == self.PAYMENT_SYNC_OPERATION:
            self.payment_sync_status_label.setText(
                f"Đồng bộ BK → Thanh toán gần nhất: {message}"
            )
        elif normalized == self.CARRIER_EXPORT_OPERATION:
            self.carrier_export_status_label.setText(f"Xuất theo bên VT gần nhất: {message}")
        else:
            self.posting_export_status_label.setText(f"Xuất kết quả nhập BK: {message}")

    def set_excel_idle(self, operation: Any | None = None) -> None:
        """Khôi phục các nút sau khi controller phát ``finished``."""

        if operation is not None:
            self._excel_operation(operation)
        self._excel_running_operation = None
        self.excel_loading_bar.set_running(False)
        for button in (
            self.sync_daily_button,
            self.post_expenses_button,
            self.sync_payment_button,
            self.export_carrier_button,
            self.export_posting_button,
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
        self.export_carrier_button.setText(
            self._excel_button_texts[self.CARRIER_EXPORT_OPERATION]
        )
        self.export_posting_button.setText(
            self._excel_button_texts[self.POSTING_EXPORT_OPERATION]
        )
        enabled = not self._rpa_running
        self.sync_daily_button.setEnabled(enabled)
        self.post_expenses_button.setEnabled(enabled)
        self.sync_payment_button.setEnabled(enabled)
        self.export_carrier_button.setEnabled(enabled)
        self.export_posting_button.setEnabled(enabled and self._posting_export_available)
        self.run_rpa_expense_button.setEnabled(enabled)

    def set_excel_actions_enabled(self, enabled: bool) -> None:
        if self._excel_running_operation is not None and enabled:
            return
        self.sync_daily_button.setEnabled(enabled)
        self.post_expenses_button.setEnabled(enabled)
        self.sync_payment_button.setEnabled(enabled)
        self.export_carrier_button.setEnabled(enabled)
        self.export_posting_button.setEnabled(enabled and self._posting_export_available)
        self.run_rpa_expense_button.setEnabled(enabled and not self._rpa_running)

    def set_rpa_running(self, message: str = "") -> None:
        self._rpa_running = True
        self.rpa_loading_bar.set_running(True)
        set_button_loading(self.run_rpa_expense_button, True)
        self.sync_daily_button.setEnabled(False)
        self.post_expenses_button.setEnabled(False)
        self.sync_payment_button.setEnabled(False)
        self.export_carrier_button.setEnabled(False)
        self.export_posting_button.setEnabled(False)
        self.run_rpa_expense_button.setEnabled(False)
        self.run_rpa_expense_button.setText("Đang chuẩn bị…")
        self.rpa_expense_status_label.setText(
            f"Nhập quyết toán: {message or 'Đang chuẩn bị dữ liệu…'}"
        )

    def set_rpa_progress(self, message: str) -> None:
        self.rpa_expense_status_label.setText(f"Nhập quyết toán: {message}")

    def set_rpa_result(self, result: Any = None) -> None:
        message = (
            result
            if isinstance(result, str)
            else _get(result, "message", default="Đã mở công cụ nhập.")
        )
        self.rpa_expense_status_label.setText(
            f"Lần nhập gần nhất: {message or 'Đã mở công cụ nhập.'}"
        )

    def set_latest_excel_data_available(
        self, operation: Any, available: bool
    ) -> None:
        normalized = self._excel_operation(operation)
        buttons = {
            self.SYNC_OPERATION: self.view_daily_sync_button,
            self.POSTING_OPERATION: self.view_expense_posting_button,
            self.PAYMENT_SYNC_OPERATION: self.view_payment_sync_button,
            self.CARRIER_EXPORT_OPERATION: self.view_carrier_export_button,
            self.POSTING_EXPORT_OPERATION: self.view_posting_export_button,
        }
        button = buttons[normalized]
        if normalized in {self.CARRIER_EXPORT_OPERATION, self.POSTING_EXPORT_OPERATION}:
            button.setText("Mở file ›" if available else "Chưa có file")
        else:
            button.setText("Xem lại ›" if available else "Chưa có dữ liệu")
        button.setEnabled(bool(available))

    def set_posting_export_available(self, available: bool, message: str = "") -> None:
        self._posting_export_available = bool(available)
        self.export_posting_button.setEnabled(
            bool(available) and not self._rpa_running and self._excel_running_operation is None
        )
        self.posting_export_status_label.setText(
            message or (
                "Sẵn sàng xuất lượt nhập gần nhất"
                if available else "Chưa có lượt nhập BK để xuất"
            )
        )

    def set_latest_rpa_data_available(self, available: bool) -> None:
        self.view_rpa_expense_button.setText(
            "Xem dữ liệu đã gửi ›" if available else "Chưa có dữ liệu đã gửi"
        )
        self.view_rpa_expense_button.setEnabled(bool(available))

    def set_rpa_idle(self) -> None:
        self._rpa_running = False
        self.rpa_loading_bar.set_running(False)
        set_button_loading(self.run_rpa_expense_button, False)
        self.run_rpa_expense_button.setText(self._rpa_button_text)
        if self._excel_running_operation is None:
            self.sync_daily_button.setEnabled(True)
            self.post_expenses_button.setEnabled(True)
            self.sync_payment_button.setEnabled(True)
            self.export_carrier_button.setEnabled(True)
            self.export_posting_button.setEnabled(self._posting_export_available)
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
