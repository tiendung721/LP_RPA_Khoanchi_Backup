"""Cửa sổ chính và lớp điều phối giữa UI với các service ứng dụng."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any

from PySide6.QtCore import QByteArray, Qt, QTimer, QUrl, Signal, Slot
from PySide6.QtGui import QCloseEvent, QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QStackedWidget,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from app.services.excel.review import CorrectionIssue, SourceDataChangedError
from app.services.excel.workbook import (
    WorkbookChangedError,
    WorkbookLockedError,
)

from .history_page import HistoryPage
from .log_page import LogPage
from .review_window import ReviewWindow
from .settings_page import SettingsPage
from .workflow_page import WorkflowPage
from .sea_freight_center import SeaFreightReconciliationDialog
from .rpa_expense_dialog import RpaLatestDataDialog, RpaSqtSelectionDialog
from .excel_dialogs import (
    ConflictResolutionDialog,
    DailySyncAllocationDialog,
    ExcelOutcomeDialog,
    MonthSelectionDialog,
    PaymentNewRowsDialog,
    PostingAllocationDialog,
    RepostSelectionDialog,
)
from .excel_summary_dialogs import (
    ExcelCompletionDialog,
    ExcelConfirmationDialog,
    daily_completion_summary,
    daily_confirmation_summary,
    payment_completion_summary,
    payment_confirmation_summary,
    posting_completion_summary,
    posting_confirmation_summary,
)

LOGGER = logging.getLogger(__name__)


@dataclass(slots=True)
class _AssistantSession:
    session_id: str
    context: str
    reconciliation_group_id: int | None = None


@dataclass(frozen=True, slots=True)
class _ExcelRetryContext:
    operation: str
    batch_id: int | None = None
    sheet_name: str | None = None
    repost_source_indices: tuple[int, ...] | None = None
    group_target_sheets: tuple[tuple[str, str], ...] | None = None
    split_document_ids: tuple[str, ...] | None = None
    highlight_unresolved: bool = False


def _attribute(source: Any, *names: str, default: Any = None) -> Any:
    if source is None:
        return default
    if isinstance(source, Mapping):
        for name in names:
            if name in source:
                return source[name]
        return default
    for name in names:
        if hasattr(source, name):
            return getattr(source, name)
    return default


def _status_code(batch: Any) -> str:
    metadata = _attribute(batch, "metadata", default=batch)
    status = _attribute(metadata, "status", default="")
    return str(getattr(status, "value", status)).split(".")[-1].upper()


def _batch_id(batch: Any) -> Any:
    metadata = _attribute(batch, "metadata", default=batch)
    return _attribute(metadata, "id", "batch_id")


class MainWindow(QMainWindow):
    """Shell điều hướng và điều phối đầy đủ cho Bước 1/Bước 2.

    ``controller`` là tùy chọn. Từng dependency cũng có thể truyền trực tiếp,
    rất thuận tiện cho ``main.py`` và unit test:

    ``MainWindow(batch_service=..., config_manager=..., assistant_launcher=...,
    watcher=..., validator=..., settings=...)``.
    """

    settingsChanged = Signal(object)
    activeBatchChanged = Signal(object)
    pageChanged = Signal(int)
    closing = Signal()

    def __init__(
        self,
        controller: Any | None = None,
        parent: QWidget | None = None,
        *,
        batch_service: Any | None = None,
        config_manager: Any | None = None,
        settings_service: Any | None = None,
        assistant_launcher: Any | None = None,
        watcher: Any | None = None,
        validator: Any | None = None,
        settings: Any | None = None,
        paths: Any | None = None,
        log_path: str | Path | None = None,
        start_watcher: bool = True,
        excel_task_controller: Any | None = None,
        excel_configuration_service: Any | None = None,
        excel_run_repository: Any | None = None,
        sea_freight_service: Any | None = None,
        rpa_expense_controller: Any | None = None,
    ) -> None:
        if isinstance(controller, QWidget) and parent is None:
            parent = controller
            controller = None
        super().__init__(parent)
        self.setObjectName("mainWindow")
        self.setWindowTitle("Trợ lý dữ liệu quyết toán")
        self.setMinimumSize(980, 650)
        self.resize(1280, 780)
        self._controller = controller
        self._batch_service = batch_service or _attribute(
            controller, "batch_service", "batches"
        )
        self._config_manager = config_manager or settings_service or _attribute(
            controller, "config_manager", "settings_service", "config"
        )
        self._assistant_launcher = assistant_launcher or _attribute(
            controller, "assistant_launcher", "launcher"
        )
        self._watcher = watcher or _attribute(controller, "watcher", "output_watcher")
        self._validator = validator or _attribute(
            controller, "validation_service", "validator"
        )
        self._excel_tasks = excel_task_controller or _attribute(
            controller, "excel_task_controller"
        )
        self._excel_configuration_service = (
            excel_configuration_service
            or _attribute(controller, "excel_configuration_service")
        )
        self._excel_run_repository = excel_run_repository or _attribute(
            controller, "excel_run_repository"
        )
        self._sea_freight_service = sea_freight_service or _attribute(
            controller, "sea_freight_service"
        )
        self._rpa_expense = (
            rpa_expense_controller
            or _attribute(controller, "rpa_expense_controller")
        )
        self._excel_operation: str | None = None
        self._excel_context: str | None = None
        self._excel_review_dialog: ConflictResolutionDialog | None = None
        self._excel_review_plan: Any | None = None
        self._excel_review_resolutions: dict[str, Any] = {}
        self._excel_review_operation: str | None = None
        self._excel_review_confirmed = False
        self._excel_reanalysis_from_source_change = False
        self._latest_excel_outcomes: dict[str, list[Any]] = {}
        self._latest_rpa_payload: dict[str, Any] | None = None
        self._paths = paths or _attribute(controller, "paths", "app_paths")
        self._settings = settings or _attribute(controller, "settings")
        self._active_batch: Any | None = None
        self._review_windows: dict[Any, ReviewWindow] = {}
        self._sea_freight_dialog: SeaFreightReconciliationDialog | None = None
        self._assistant_sessions: list[_AssistantSession] = []
        self._closing = False
        self._watcher_connected = False

        self._ensure_default_helpers()
        self._load_settings_if_needed()
        if self._paths is None:
            self._paths = _attribute(self._settings, "paths")
        resolved_log_path = log_path or _attribute(self._paths, "log_path")

        self._build_ui(resolved_log_path)
        self._connect_page_signals()
        self._connect_service_signals()
        self._load_initial_data()
        self._restore_ui_state()
        if start_watcher:
            self._start_watcher()

    def _ensure_default_helpers(self) -> None:
        if self._validator is None:
            try:
                from app.services.validation_service import ValidationService

                self._validator = ValidationService()
            except (ImportError, TypeError):
                self._validator = None
        if self._assistant_launcher is None:
            try:
                from app.services.assistant_bat_launcher import AssistantBatLauncher

                self._assistant_launcher = AssistantBatLauncher(self._settings)
            except (ImportError, TypeError):
                self._assistant_launcher = None

    def _load_settings_if_needed(self) -> None:
        if self._settings is not None:
            return
        owner = self._config_manager or self._controller
        for name in ("load", "load_settings", "get_settings"):
            method = getattr(owner, name, None) if owner is not None else None
            if callable(method):
                try:
                    self._settings = method()
                except Exception as exc:
                    LOGGER.exception("Không thể nạp cấu hình: %s", exc)
                break

    def _build_ui(self, log_path: str | Path | None) -> None:
        root = QWidget()
        root.setObjectName("applicationRoot")
        self.setCentralWidget(root)
        shell = QHBoxLayout(root)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)

        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(232)
        sidebar.setStyleSheet(
            "QFrame#sidebar {"
            "background: qlineargradient(x1:0, y1:0, x2:0, y2:1, "
            "stop:0 #142B4A, stop:1 #0E1F37);"
            "}"
        )
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(13, 20, 13, 15)
        brand = QLabel("TRỢ LÝ DỮ LIỆU\nQUYẾT TOÁN")
        brand.setStyleSheet(
            "color: white; font-size: 13pt; font-weight: 700; "
            "letter-spacing: 0.5px; padding: 4px 9px 16px 9px;"
        )
        sidebar_layout.addWidget(brand)
        self.navigation = QListWidget()
        self.navigation.setObjectName("navigation")
        self.navigation.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        for text in ("Thao tác", "Lịch sử", "Cài đặt", "Nhật ký"):
            item = QListWidgetItem(text)
            item.setTextAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
            self.navigation.addItem(item)
        sidebar_layout.addWidget(self.navigation, 1)
        self.watcher_status = QLabel("Đang khởi tạo bộ theo dõi…")
        self.watcher_status.setObjectName("watcherStatus")
        self.watcher_status.setWordWrap(True)
        self.watcher_status.setStyleSheet(
            "color: #C7D5E8; background: rgba(255, 255, 255, 0.07); "
            "border: 1px solid rgba(255, 255, 255, 0.08); "
            "border-radius: 8px; padding: 9px;"
        )
        sidebar_layout.addWidget(self.watcher_status)
        shell.addWidget(sidebar)

        self.pages = QStackedWidget()
        self.pages.setObjectName("mainPages")
        self.workflow_page = WorkflowPage(self._settings)
        self.history_page = HistoryPage()
        self.settings_page = SettingsPage(self._settings)
        self.log_page = LogPage(log_path)
        self.pages.addWidget(self.workflow_page)
        self.pages.addWidget(self.history_page)
        self.pages.addWidget(self.settings_page)
        self.pages.addWidget(self.log_page)
        shell.addWidget(self.pages, 1)

        status = QStatusBar()
        self.setStatusBar(status)
        self.statusBar().showMessage("Sẵn sàng")
        self.navigation.setCurrentRow(0)

    def _connect_page_signals(self) -> None:
        self.navigation.currentRowChanged.connect(self._change_page)
        self.workflow_page.open_assistant_requested.connect(self.open_assistant)
        self.workflow_page.open_bang_ke_assistant_requested.connect(
            self.open_bang_ke_assistant
        )
        self.workflow_page.open_review_requested.connect(self.open_review)
        self.workflow_page.sync_daily_requested.connect(self.start_daily_sync)
        self.workflow_page.post_expenses_requested.connect(
            self.start_expense_posting
        )
        self.workflow_page.sync_payment_requested.connect(
            self.start_payment_sync
        )
        self.workflow_page.run_rpa_expense_requested.connect(
            self.start_rpa_expense
        )
        self.workflow_page.view_latest_excel_requested.connect(
            self.show_latest_excel_data
        )
        self.workflow_page.view_latest_rpa_requested.connect(
            self.show_latest_rpa_data
        )
        self.history_page.refresh_requested.connect(self.refresh_history)
        self.history_page.open_batch_requested.connect(self.open_review)
        self.history_page.open_path_requested.connect(self.open_containing_folder)

        self.settings_page.save_requested.connect(self.save_settings)
        self.settings_page.check_requested.connect(self.check_settings)
        self.settings_page.open_output_requested.connect(self._open_directory)

    def _connect_service_signals(self) -> None:
        watcher = self._watcher
        if watcher is not None and not self._watcher_connected:
            self._safe_connect(watcher, "file_rejected", self._watcher_rejected)
            self._safe_connect(watcher, "watcher_error", self._watcher_error)
            self._safe_connect(watcher, "status_changed", self._watcher_status_changed)
            self._safe_connect(watcher, "started", lambda path: self._watcher_status_changed(True, f"Đang theo dõi: {path}"))
            self._safe_connect(watcher, "stopped", lambda: self._watcher_status_changed(False, "Đã dừng theo dõi Output."))
            self._safe_connect(watcher, "scan_completed", self._scan_completed)
            callback = getattr(watcher, "_on_file_ready_callback", None)
            if callback is not None and hasattr(watcher, "file_processed"):
                self._safe_connect(watcher, "file_processed", self._file_processed)
            else:
                self._safe_connect(watcher, "file_ready", self._file_ready)
                self._safe_connect(watcher, "file_processed", self._file_processed)
            self._watcher_connected = True

        service = self._batch_service
        if service is not None:
            for signal_name in ("batch_received", "batch_created", "batchUpdated"):
                self._safe_connect(service, signal_name, self._external_batch_changed)
            for signal_name in ("active_batch_changed", "activeBatchChanged"):
                self._safe_connect(service, signal_name, self._external_batch_changed)

        excel_tasks = self._excel_tasks
        if excel_tasks is not None:
            self._safe_connect(excel_tasks, "started", self._excel_started)
            self._safe_connect(excel_tasks, "progress", self._excel_progress)
            self._safe_connect(
                excel_tasks, "analysis_ready", self._excel_analysis_ready
            )
            self._safe_connect(
                excel_tasks,
                "correction_required",
                self._excel_correction_required,
            )
            self._safe_connect(excel_tasks, "completed", self._excel_completed)
            self._safe_connect(excel_tasks, "failed", self._excel_failed)
            self._safe_connect(excel_tasks, "finished", self._excel_finished)

        rpa = self._rpa_expense
        if rpa is not None:
            self._safe_connect(rpa, "started", self._rpa_started)
            self._safe_connect(rpa, "progress", self._rpa_progress)
            self._safe_connect(rpa, "sheets_ready", self._rpa_sheets_ready)
            self._safe_connect(rpa, "plan_ready", self._rpa_plan_ready)
            self._safe_connect(rpa, "launched", self._rpa_launched)
            self._safe_connect(rpa, "failed", self._rpa_failed)
            self._safe_connect(rpa, "finished", self._rpa_finished)

    @staticmethod
    def _safe_connect(owner: Any, signal_name: str, slot: Callable[..., Any]) -> bool:
        signal = getattr(owner, signal_name, None)
        connect = getattr(signal, "connect", None)
        if callable(connect):
            connect(slot)
            return True
        return False

    def _load_initial_data(self) -> None:
        self.workflow_page.set_configuration(self._settings)
        self.settings_page.set_settings(self._settings)
        self._load_excel_history()
        self._load_rpa_history()
        self.refresh_history(silent=True)
        if self._batch_service is None:
            self.workflow_page.clear_active_batch()
            return
        current_output = getattr(
            self._batch_service, "get_current_output_batch", None
        )
        if callable(current_output):
            try:
                batch = current_output()
            except Exception as exc:
                LOGGER.exception(
                    "Không thể nhận diện batch Output hiện hành: %s", exc
                )
            else:
                if batch is not None:
                    self._set_active_batch(batch)
                else:
                    self.workflow_page.clear_active_batch()
                return
        for name in ("restore_active_batch", "get_active_batch", "restore_last_batch"):
            method = getattr(self._batch_service, name, None)
            if not callable(method):
                continue
            try:
                review = method()
                if review is not None:
                    self._set_active_batch(review)
            except Exception as exc:
                LOGGER.exception("Không thể khôi phục batch đang làm dở: %s", exc)
                self.statusBar().showMessage(
                    "Không khôi phục được batch đang làm dở; hãy chọn trong Lịch sử.",
                    8000,
                )
            break

    def _restore_ui_state(self) -> None:
        owner = self._controller
        if owner is None:
            return
        state: Any = None
        for name in ("restore_ui_state", "load_ui_state", "get_ui_state"):
            method = getattr(owner, name, None)
            if callable(method):
                try:
                    state = method()
                except Exception:
                    LOGGER.exception("Không thể khôi phục trạng thái cửa sổ.")
                break
        if not isinstance(state, Mapping):
            return
        geometry = state.get("geometry")
        if geometry:
            try:
                if isinstance(geometry, str):
                    geometry = QByteArray.fromBase64(geometry.encode("ascii"))
                self.restoreGeometry(geometry)
            except (TypeError, ValueError):
                LOGGER.warning("Geometry đã lưu không hợp lệ.")
        page = state.get("page", state.get("page_index"))
        if isinstance(page, int) and 0 <= page < self.pages.count():
            self.navigation.setCurrentRow(page)

    @Slot(int)
    def _change_page(self, index: int) -> None:
        if not 0 <= index < self.pages.count():
            return
        self.pages.setCurrentIndex(index)
        self.pageChanged.emit(index)
        if index == 1:
            self.refresh_history(silent=True)
        elif index == 3:
            self.log_page.refresh()

    @Slot()
    def open_assistant(self) -> None:
        self._launch_assistant(self._settings, context="home")

    @Slot()
    def open_bang_ke_assistant(self) -> None:
        self._launch_assistant(
            self._settings,
            context="bang_ke",
            bat_setting="bang_ke_assistant_bat_path",
        )

    def open_reconciliation_assistant(self, group_id: int) -> None:
        self._launch_assistant(
            self._settings,
            context="reconciliation",
            reconciliation_group_id=group_id,
        )

    def _missing_excel_configuration(
        self,
        *,
        require_daily: bool,
        require_payment: bool = False,
    ) -> bool:
        daily = str(
            _attribute(self._settings, "daily_workbook_path", default="") or ""
        ).strip()
        bk = str(
            _attribute(self._settings, "bk_workbook_path", default="") or ""
        ).strip()
        payment = str(
            _attribute(
                self._settings,
                "payment_workbook_path",
                default="",
            )
            or ""
        ).strip()
        if (
            bk
            and (daily or not require_daily)
            and (payment or not require_payment)
        ):
            return False
        message = QMessageBox(self)
        message.setIcon(QMessageBox.Icon.Warning)
        message.setWindowTitle("Chưa cấu hình đường dẫn Excel")
        message.setText(
            "Hãy cấu hình file Hàng ngày và file BK trong trang Cài đặt."
            if require_daily
            else "Hãy cấu hình file BK trong trang Cài đặt."
        )
        open_settings = message.addButton(
            "Mở Cài đặt", QMessageBox.ButtonRole.AcceptRole
        )
        message.addButton("Đóng", QMessageBox.ButtonRole.RejectRole)
        message.exec()
        if message.clickedButton() is open_settings:
            self.navigation.setCurrentRow(2)
        return True

    @Slot()
    def start_daily_sync(self) -> None:
        if self._missing_excel_configuration(require_daily=True):
            return
        if self._excel_tasks is None:
            QMessageBox.warning(
                self,
                "Chưa thể đồng bộ",
                "Dịch vụ xử lý Excel chưa được khởi tạo.",
            )
            return
        self._excel_context = "workflow"
        try:
            service = _attribute(self._excel_tasks, "daily_sync_service")
            list_candidates = getattr(service, "source_sheet_candidates", None)
            if not callable(list_candidates):
                raise RuntimeError(
                    "Dịch vụ đồng bộ chưa hỗ trợ chọn sheet nguồn."
                )
            candidates = list(list_candidates())
            dialog = MonthSelectionDialog(
                candidates,
                self,
                title="Chọn sheet tháng cần đồng bộ",
                preselect_first=False,
                show_recommendations=False,
                multi_select=True,
            )
            if dialog.exec() != QDialog.DialogCode.Accepted:
                self._excel_context = None
                return
            source_sheet_names = dialog.selected_sheet_names
            if not source_sheet_names:
                self._excel_context = None
                return
            self._excel_tasks.start_sync(
                source_sheet_names=source_sheet_names
            )
        except Exception as exc:
            self._show_excel_error(exc, operation="sync")

    @Slot()
    def start_expense_posting(self) -> None:
        if self._missing_excel_configuration(require_daily=False):
            return
        if self._excel_tasks is None:
            QMessageBox.warning(
                self,
                "Chưa thể nhập khoản chi",
                "Dịch vụ xử lý Excel chưa được khởi tạo.",
            )
            return
        self._excel_context = "workflow"
        try:
            batch_id = _batch_id(self._active_batch)
            if batch_id is None:
                raise RuntimeError(
                    "Chưa có file bóc tách hiện hành đã xác nhận để nhập khoản chi."
                )
            self._excel_tasks.start_posting(batch_id=int(batch_id))
        except Exception as exc:
            self._show_excel_error(exc, operation="posting")

    @Slot()
    def start_payment_sync(self) -> None:
        if self._missing_excel_configuration(
            require_daily=False,
            require_payment=True,
        ):
            return
        if self._excel_tasks is None:
            QMessageBox.warning(
                self,
                "Chưa thể đồng bộ",
                "Dịch vụ đồng bộ BK sang Thanh toán chưa được khởi tạo.",
            )
            return
        self._excel_context = "workflow"
        try:
            service = _attribute(self._excel_tasks, "payment_sync_service")
            list_candidates = getattr(service, "source_sheet_candidates", None)
            if not callable(list_candidates):
                raise RuntimeError(
                    "Dịch vụ đồng bộ Thanh toán chưa hỗ trợ chọn sheet BK."
                )
            candidates = list(list_candidates())
            dialog = MonthSelectionDialog(
                candidates,
                self,
                title="Chọn sheet BK đồng bộ sang Thanh toán",
                preselect_first=False,
                show_recommendations=False,
                multi_select=True,
            )
            if dialog.exec() != QDialog.DialogCode.Accepted:
                self._excel_context = None
                return
            source_sheet_names = dialog.selected_sheet_names
            if not source_sheet_names:
                self._excel_context = None
                return
            self._excel_tasks.start_payment_sync(
                source_sheet_names=source_sheet_names
            )
        except Exception as exc:
            self._show_excel_error(exc, operation="payment_sync")

    def _missing_rpa_configuration(self) -> bool:
        bk = str(
            _attribute(self._settings, "bk_workbook_path", default="") or ""
        ).strip()
        bat = str(
            _attribute(
                self._settings,
                "rpa_expense_bat_path",
                default="",
            )
            or ""
        ).strip()
        if bk and bat:
            return False
        message = QMessageBox(self)
        message.setIcon(QMessageBox.Icon.Warning)
        message.setWindowTitle("Chưa cấu hình luồng RPA")
        message.setText(
            "Hãy cấu hình file BK và BAT RPA nhập quyết toán trong trang Cài đặt."
        )
        open_settings = message.addButton(
            "Mở Cài đặt", QMessageBox.ButtonRole.AcceptRole
        )
        message.addButton("Đóng", QMessageBox.ButtonRole.RejectRole)
        message.exec()
        if message.clickedButton() is open_settings:
            self.navigation.setCurrentRow(2)
        return True

    @Slot()
    def start_rpa_expense(self) -> None:
        if self._missing_rpa_configuration():
            return
        if self._rpa_expense is None:
            QMessageBox.warning(
                self,
                "Chưa thể chạy RPA",
                "Dịch vụ chuẩn bị dữ liệu RPA chưa được khởi tạo.",
            )
            return
        if self._excel_tasks is not None and self._excel_tasks.is_busy:
            QMessageBox.information(
                self,
                "Excel đang được sử dụng",
                "Hãy chờ tác vụ Excel hiện tại hoàn tất rồi chạy RPA.",
            )
            return
        try:
            self._rpa_expense.load_sheets()
        except Exception as exc:
            self._rpa_failed(exc)

    @Slot(str)
    def _rpa_started(self, phase: str) -> None:
        labels = {
            "sheets": "Đang đọc danh sách sheet BK…",
            "analysis": "Đang tổng hợp dữ liệu theo SQT…",
            "launch": "Đang tạo dữ liệu và khởi chạy PAD…",
        }
        message = labels.get(phase, "Đang chuẩn bị dữ liệu RPA…")
        self.workflow_page.set_rpa_running(message)
        self.statusBar().showMessage(message)

    @Slot(str)
    def _rpa_progress(self, message: str) -> None:
        self.workflow_page.set_rpa_progress(message)
        self.statusBar().showMessage(message)

    @Slot(object)
    def _rpa_sheets_ready(self, candidates: Any) -> None:
        dialog = MonthSelectionDialog(
            list(candidates),
            self,
            title="Chọn sheet BK chạy RPA nhập quyết toán",
            preselect_first=False,
            show_recommendations=False,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self.workflow_page.set_rpa_result("Đã hủy trước khi chọn sheet.")
            return
        sheet_name = dialog.selected_sheet_name
        if not sheet_name:
            return
        try:
            self._rpa_expense.analyze_sheet(sheet_name)
        except Exception as exc:
            self._rpa_failed(exc)

    @Slot(object)
    def _rpa_plan_ready(self, plan: Any) -> None:
        dialog = RpaSqtSelectionDialog(plan, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self.workflow_page.set_rpa_result("Đã hủy trước khi chạy PAD.")
            return
        try:
            self._rpa_expense.launch(plan, dialog.selected_sqt)
        except Exception as exc:
            self._rpa_failed(exc)

    @Slot(object)
    def _rpa_launched(self, result: Any) -> None:
        self.workflow_page.set_rpa_result(result)
        self._load_rpa_history()
        self.statusBar().showMessage(
            str(_attribute(result, "message", default="Đã khởi chạy PAD.")),
            10000,
        )
        message = QMessageBox(self)
        message.setIcon(QMessageBox.Icon.Information)
        message.setWindowTitle("Đã khởi chạy RPA")
        message.setText(
            f"{_attribute(result, 'message', default='Đã khởi chạy PAD.')}\n\n"
            f"File dữ liệu: {_attribute(result, 'selection_path', default='')}\n"
            "PAD chỉ đánh dấu “Đã nhập” sau khi thao tác Lưu trên web thành công.",
        )
        view_details = message.addButton(
            "Xem dữ liệu đã gửi", QMessageBox.ButtonRole.ActionRole
        )
        message.addButton(QMessageBox.StandardButton.Ok)
        message.exec()
        if message.clickedButton() is view_details:
            self.show_latest_rpa_data()

    @Slot(object)
    def _rpa_failed(self, error: Any) -> None:
        LOGGER.error("Luồng RPA thất bại: %s", error)
        self.workflow_page.set_rpa_result(f"Lỗi: {error}")
        QMessageBox.critical(
            self,
            "Không thể chạy RPA",
            f"{error}\n\nHãy kiểm tra file BK đã đóng và cấu hình BAT RPA.",
        )

    @Slot(str)
    def _rpa_finished(self, _phase: str) -> None:
        self.workflow_page.set_rpa_idle()
        self.statusBar().showMessage("Sẵn sàng", 3000)

    @Slot(str)
    def _excel_started(self, operation: str) -> None:
        self._excel_operation = operation
        if self._excel_context != "configuration":
            self.workflow_page.set_excel_running(operation)
        self.statusBar().showMessage("Đang xử lý dữ liệu Excel…")

    @Slot(str, str)
    def _excel_progress(self, operation: str, message: str) -> None:
        if self._excel_context == "configuration":
            self.settings_page.show_check_result(True, message)
        else:
            self.workflow_page.set_excel_progress(operation, message)
        self.statusBar().showMessage(message)

    def _saved_excel_resolutions(
        self, plan: Any, operation: str
    ) -> dict[str, Any]:
        loader = getattr(self._excel_tasks, "saved_resolutions", None)
        if not callable(loader):
            self._excel_restore_info = {}
            return {}
        resolutions = dict(loader(plan, operation=operation) or {})
        info = getattr(self._excel_tasks, "saved_resolution_info", {})
        self._excel_restore_info = dict(info or {})
        return resolutions

    def _save_excel_draft(
        self,
        plan: Any,
        resolutions: Mapping[str, Any],
        operation: str,
    ) -> None:
        saver = getattr(self._excel_tasks, "save_draft", None)
        if callable(saver):
            saver(plan, resolutions, operation=operation)

    @Slot(object)
    def _excel_analysis_ready(self, plan: Any) -> None:
        if self._excel_tasks is None:
            return
        try:
            operation = self._excel_tasks.normalize_operation(
                _attribute(plan, "operation", "operation_type", default=self._excel_operation)
            )
            if operation == "payment_sync":
                self._handle_payment_sync_plan(plan)
                return
            resolutions = MainWindow._saved_excel_resolutions(
                self, plan, operation
            )
            source_reload_issues: list[CorrectionIssue] = []
            handled_conflicts: set[str] = set()
            conflicts = list(_attribute(plan, "conflicts", default=()) or ())
            selected_sync_sheet = _attribute(
                plan, "selected_sheet", "selected_sheet_name"
            )

            selected_month = _attribute(plan, "selected_month")
            month_candidates = list(
                _attribute(plan, "month_candidates", default=()) or ()
            )
            source_target_sheets = dict(
                _attribute(plan, "source_target_sheets", default={}) or {}
            )
            if (
                operation == "sync"
                and source_target_sheets
                and any(target in (None, "") for target in source_target_sheets.values())
            ):
                dialog = DailySyncAllocationDialog(plan, self)
                if dialog.exec() != QDialog.DialogCode.Accepted:
                    self._excel_tasks.cancel_waiting()
                    return
                self._excel_tasks.cancel_waiting()
                self._excel_context = "workflow"
                self._excel_tasks.start_sync(
                    source_sheet_names=list(source_target_sheets),
                    source_target_sheets=dialog.source_target_sheets,
                )
                return
            if (
                operation == "sync"
                and not source_target_sheets
                and selected_month is None
                and month_candidates
            ):
                dialog = MonthSelectionDialog(
                    month_candidates,
                    self,
                    title="Chọn tháng cần đồng bộ",
                )
                if dialog.exec() != QDialog.DialogCode.Accepted:
                    self._excel_tasks.cancel_waiting()
                    return
                selection = dialog.selection()
                selected_sync_sheet = selection.get(
                    "selected_sheet_name",
                    selection.get("selected_sheet"),
                )
                for conflict in conflicts:
                    kind = str(
                        getattr(
                            _attribute(conflict, "conflict_type", "type", default=""),
                            "value",
                            _attribute(conflict, "conflict_type", "type", default=""),
                        )
                    )
                    if kind == "TARGET_MONTH_AMBIGUOUS":
                        conflict_id = str(
                            _attribute(conflict, "conflict_id", "id")
                        )
                        resolutions[conflict_id] = {
                            "conflict_id": conflict_id,
                            "action": "SELECT_MONTH",
                            **selection,
                        }
                        handled_conflicts.add(conflict_id)

            selected_sheet = _attribute(plan, "selected_sheet", "selected_sheet_name")
            raw_source_kind = _attribute(plan, "source_kind", default="ASSISTANT")
            source_kind = str(
                getattr(raw_source_kind, "value", raw_source_kind) or "ASSISTANT"
            ).strip().upper()
            sheet_candidates = list(
                _attribute(
                    plan,
                    "sheet_candidates",
                    "target_sheet_candidates",
                    default=(),
                )
                or ()
            )
            source_groups = list(
                _attribute(plan, "source_groups", default=()) or ()
            )
            restored_splits = {
                str(value)
                for value in (
                    resolutions.pop("split_document_ids", ()) or ()
                )
                if str(value).strip()
            }
            current_splits = {
                str(value)
                for value in (
                    _attribute(plan, "split_document_ids", default=()) or ()
                )
            }
            if (
                operation == "posting"
                and source_kind != "BANG_KE"
                and restored_splits.difference(current_splits)
            ):
                self._excel_tasks.cancel_waiting()
                self._excel_context = "workflow"
                self._excel_tasks.start_posting(
                    batch_id=_attribute(plan, "batch_id", default=None),
                    split_document_ids=sorted(current_splits | restored_splits),
                )
                return
            restored_group_targets = resolutions.pop("group_target_sheets", {})
            if source_groups and isinstance(restored_group_targets, Mapping):
                for group in source_groups:
                    group_id = str(_attribute(group, "group_id", default=""))
                    restored_target = restored_group_targets.get(group_id)
                    if (
                        restored_target not in (None, "")
                        and not bool(_attribute(group, "target_locked", default=False))
                    ):
                        setattr(group, "target_sheet", str(restored_target))
            if (
                operation == "posting"
                and source_kind != "BANG_KE"
                and source_groups
                and any(
                    _attribute(group, "target_sheet", default=None) in (None, "")
                    for group in source_groups
                )
            ):
                dialog = PostingAllocationDialog(plan, self)
                if dialog.exec() != QDialog.DialogCode.Accepted:
                    self._excel_tasks.cancel_waiting()
                    return
                if dialog.split_document_ids:
                    self._excel_tasks.cancel_waiting()
                    self._excel_context = "workflow"
                    self._excel_tasks.start_posting(
                        batch_id=_attribute(plan, "batch_id", default=None),
                        split_document_ids=sorted(
                            current_splits | dialog.split_document_ids
                        ),
                    )
                    return
                self._excel_tasks.cancel_waiting()
                self._excel_context = "workflow"
                self._excel_tasks.start_posting(
                    batch_id=_attribute(plan, "batch_id", default=None),
                    group_target_sheets=dialog.group_target_sheets,
                    split_document_ids=sorted(current_splits),
                )
                return
            if (
                operation == "posting"
                and source_kind != "BANG_KE"
                and source_groups
                and restored_group_targets
                and all(
                    _attribute(group, "target_sheet", default=None) not in (None, "")
                    for group in source_groups
                )
                and not list(_attribute(plan, "items", default=()) or ())
            ):
                self._excel_tasks.cancel_waiting()
                self._excel_context = "workflow"
                self._excel_tasks.start_posting(
                    batch_id=_attribute(plan, "batch_id", default=None),
                    group_target_sheets={
                        str(_attribute(group, "group_id")): str(
                            _attribute(group, "target_sheet")
                        )
                        for group in source_groups
                    },
                )
                return
            if (
                operation == "posting"
                and source_kind != "BANG_KE"
                and not source_groups
                and selected_sheet in (None, "")
                and sheet_candidates
            ):
                dialog = MonthSelectionDialog(
                    sheet_candidates,
                    self,
                    title="Chọn sheet nhận khoản chi",
                    preselect_first=False,
                    show_recommendations=False,
                )
                if dialog.exec() != QDialog.DialogCode.Accepted:
                    self._excel_tasks.cancel_waiting()
                    return
                # Re-analyze the selected sheet so every row/cell conflict is
                # collected before apply; compatible choices are restored after
                # the selected sheet has been analyzed.
                sheet_name = dialog.selected_sheet_name
                self._excel_tasks.cancel_waiting()
                self._excel_context = "workflow"
                self._excel_tasks.start_posting(
                    batch_id=_attribute(plan, "batch_id", default=None),
                    sheet_name=sheet_name,
                )
                return

            previously_posted = list(
                _attribute(plan, "previously_posted_items", default=()) or ()
            )
            repost_selection_done = bool(
                _attribute(plan, "repost_selection_done", default=False)
            )
            if (
                operation == "posting"
                and (
                    selected_sheet not in (None, "")
                    or (
                        source_groups
                        and all(
                            _attribute(group, "target_sheet", default=None)
                            not in (None, "")
                            for group in source_groups
                        )
                    )
                    or source_kind == "BANG_KE"
                )
                and previously_posted
                and not repost_selection_done
            ):
                dialog = RepostSelectionDialog(previously_posted, self)
                if dialog.exec() != QDialog.DialogCode.Accepted:
                    self._excel_tasks.cancel_waiting()
                    return
                repost_indices = dialog.selected_source_indices
                self._excel_tasks.cancel_waiting()
                self._excel_context = "workflow"
                analyze_kwargs: dict[str, Any] = {
                    "batch_id": _attribute(plan, "batch_id", default=None),
                    "repost_source_indices": repost_indices,
                }
                if source_kind != "BANG_KE":
                    if source_groups:
                        analyze_kwargs["group_target_sheets"] = {
                            str(_attribute(group, "group_id")): str(
                                _attribute(group, "target_sheet")
                            )
                            for group in source_groups
                            if _attribute(group, "target_sheet") not in (None, "")
                        }
                        analyze_kwargs["split_document_ids"] = sorted(
                            _attribute(plan, "split_document_ids", default=()) or ()
                        )
                    else:
                        analyze_kwargs["sheet_name"] = str(selected_sheet)
                self._excel_tasks.start_posting(
                    **analyze_kwargs,
                )
                return

            remaining = [
                conflict
                for conflict in conflicts
                if str(_attribute(conflict, "conflict_id", "id"))
                not in handled_conflicts
                and not (
                    operation == "sync"
                    and selected_sync_sheet not in (None, "")
                    and str(
                        _attribute(
                            _attribute(conflict, "details", default={}),
                            "target_sheet",
                            default="",
                        )
                    )
                    not in ("", str(selected_sync_sheet))
                )
            ]
            blocking_sync = [
                conflict
                for conflict in remaining
                if str(
                    getattr(
                        _attribute(
                            conflict,
                            "conflict_type",
                            "type",
                            default="",
                        ),
                        "value",
                        _attribute(
                            conflict,
                            "conflict_type",
                            "type",
                            default="",
                        ),
                    )
                )
                == "SYNC_GROUP_COUNT_MISMATCH"
            ]
            if blocking_sync:
                detail = "\n".join(
                    f"• {_attribute(conflict, 'message', default='')}"
                    for conflict in blocking_sync
                )
                QMessageBox.warning(
                    self,
                    "Không thể đồng bộ",
                    "Số dòng của cùng một SQT không khớp giữa nguồn và BK.\n"
                    "Hãy sửa dữ liệu rồi chạy lại.\n\n"
                    + detail,
                )
                self._excel_tasks.cancel_waiting()
                return
            remaining = [
                conflict for conflict in remaining if conflict not in blocking_sync
            ]
            if bool(
                getattr(self, "_excel_reanalysis_from_source_change", False)
            ):
                source_reload_issues = [
                    CorrectionIssue.from_conflict(
                        conflict,
                        issue_id=(
                            "source-reload:"
                            + str(_attribute(conflict, "conflict_id", "id", default=""))
                        ),
                        message=(
                            "Xung đột cần kiểm tra lại sau khi dữ liệu JSON thay đổi."
                        ),
                    )
                    for conflict in remaining
                    if str(_attribute(conflict, "conflict_id", "id", default=""))
                    not in resolutions
                ]
            if remaining:
                if getattr(self, "_excel_review_dialog", None) is None:
                    dialog = ConflictResolutionDialog(
                        remaining,
                        self,
                        operation=operation,
                        initial_resolutions=resolutions,
                        restore_info=getattr(self, "_excel_restore_info", {}),
                        issues=source_reload_issues,
                    )
                    setattr(self, "_excel_review_dialog", dialog)
                    setattr(self, "_excel_review_plan", plan)
                    setattr(self, "_excel_review_operation", operation)
                else:
                    dialog = getattr(self, "_excel_review_dialog")
                    dialog.set_review(
                        remaining,
                        issues=source_reload_issues,
                        initial_resolutions={
                            **getattr(self, "_excel_review_resolutions", {}),
                            **resolutions,
                        },
                    )
                setattr(self, "_excel_reanalysis_from_source_change", False)
                if dialog.exec() != QDialog.DialogCode.Accepted:
                    self._excel_tasks.cancel_waiting()
                    MainWindow._clear_excel_review(self)
                    return
                resolutions.update(dialog.resolution_map())
                review_resolutions = getattr(
                    self, "_excel_review_resolutions", {}
                )
                review_resolutions.update(resolutions)
                setattr(self, "_excel_review_resolutions", review_resolutions)
                MainWindow._save_excel_draft(
                    self, plan, resolutions, operation
                )
            if (
                operation == "posting"
                and not remaining
                and bool(_attribute(plan, "confirmation_required", default=False))
                and not bool(_attribute(plan, "confirmation_done", default=False))
            ):
                summary = posting_confirmation_summary(plan)
                dialog = ExcelConfirmationDialog(
                    summary,
                    confirm_label=(
                        f"Nhập {summary.change_count} khoản vào BK"
                        if summary.change_count
                        else "Tiếp tục kiểm tra BK"
                    ),
                    parent=self,
                )
                if dialog.exec() != QDialog.DialogCode.Accepted:
                    self._excel_tasks.cancel_waiting()
                    return
                plan.confirmation_done = True
                if getattr(self, "_excel_review_plan", None) is not None:
                    setattr(self, "_excel_review_confirmed", True)
            if operation == "sync" and not bool(
                getattr(self, "_excel_review_confirmed", False)
            ):
                selected_sync_targets = {
                    str(target)
                    for target in source_target_sheets.values()
                    if target not in (None, "")
                }
                selected_candidates = [
                    item
                    for item in month_candidates
                    if str(
                        _attribute(item, "target_sheet", "sheet_name", default="")
                    ) in selected_sync_targets
                ]
                candidate = next(
                    (
                        item
                        for item in month_candidates
                        if str(
                            _attribute(
                                item,
                                "target_sheet",
                                "sheet_name",
                                default="",
                            )
                        )
                        == str(selected_sync_sheet)
                    ),
                    None,
                )
                if selected_candidates:
                    selected_sync_sheet = ", ".join(
                        str(_attribute(item, "target_sheet", "sheet_name"))
                        for item in selected_candidates
                    )
                display_candidates = selected_candidates or (
                    [candidate] if candidate is not None else []
                )
                summary = daily_confirmation_summary(
                    plan,
                    candidates=display_candidates,
                )
                # Khi không có thay đổi, service vẫn cần chạy để ghi nhận kết
                # quả NO_CHANGES nhưng không cần bắt người dùng xác nhận ghi file.
                if summary.change_count:
                    dialog = ExcelConfirmationDialog(
                        summary,
                        confirm_label=(
                            f"Đồng bộ {summary.change_count} thay đổi vào BK"
                        ),
                        parent=self,
                    )
                    if dialog.exec() != QDialog.DialogCode.Accepted:
                        self._excel_tasks.cancel_waiting()
                        return
                if getattr(self, "_excel_review_plan", None) is not None:
                    setattr(self, "_excel_review_confirmed", True)
            selector_actions = {
                str(value.get("action", ""))
                for value in resolutions.values()
                if isinstance(value, Mapping)
            }
            selector_refine_required = bool(
                selector_actions.intersection(
                    {
                        "SELECT_SHEET",
                        "SELECT_ROW",
                        "SELECT_FEE",
                        "SELECT_INVOICE",
                        "SELECT_SOURCE_ITEM",
                        "SELECT_CARRIER",
                    }
                )
            )
            bang_ke_refine_required = source_kind == "BANG_KE" and bool(conflicts)
            review_prepare_required = bool(remaining)
            if operation == "posting" and not remaining:
                setattr(self, "_excel_reanalysis_from_source_change", False)
            if review_prepare_required or (
                operation == "posting"
                and (selector_refine_required or bang_ke_refine_required)
            ):
                prepare = getattr(self._excel_tasks, "prepare_plan", None)
                if not callable(prepare):
                    prepare = getattr(self._excel_tasks, "refine_plan", None)
                if not callable(prepare):
                    prepare = self._excel_tasks.apply_plan
                prepare(
                    plan,
                    resolutions,
                    operation=operation,
                )
            else:
                self._excel_tasks.apply_plan(
                    plan,
                    resolutions,
                    operation=operation,
                )
        except Exception as exc:
            LOGGER.exception("Không xử lý được kế hoạch Excel: %s", exc)
            self._excel_tasks.cancel_waiting()
            self._show_excel_error(exc, operation=self._excel_operation)

    def _handle_payment_sync_plan(self, plan: Any) -> None:
        """Thu thập lựa chọn rồi xác nhận một lần trước khi ghi."""

        if self._excel_tasks is None:
            return
        resolutions = MainWindow._saved_excel_resolutions(
            self, plan, "payment_sync"
        )
        conflicts = list(_attribute(plan, "conflicts", default=()) or ())
        if conflicts:
            if getattr(self, "_excel_review_dialog", None) is None:
                conflict_dialog = ConflictResolutionDialog(
                    conflicts,
                    self,
                    operation="payment_sync",
                    initial_resolutions=resolutions,
                    restore_info=getattr(self, "_excel_restore_info", {}),
                )
                setattr(self, "_excel_review_dialog", conflict_dialog)
                setattr(self, "_excel_review_plan", plan)
                setattr(self, "_excel_review_operation", "payment_sync")
            else:
                conflict_dialog = getattr(self, "_excel_review_dialog")
                conflict_dialog.set_review(
                    conflicts,
                    initial_resolutions={
                        **getattr(self, "_excel_review_resolutions", {}),
                        **resolutions,
                    },
                )
            if conflict_dialog.exec() != QDialog.DialogCode.Accepted:
                self._excel_tasks.cancel_waiting()
                MainWindow._clear_excel_review(self)
                return
            resolutions.update(conflict_dialog.resolution_map())
            review_resolutions = getattr(
                self, "_excel_review_resolutions", {}
            )
            review_resolutions.update(resolutions)
            setattr(self, "_excel_review_resolutions", review_resolutions)
            MainWindow._save_excel_draft(
                self, plan, resolutions, "payment_sync"
            )
            prepare = getattr(self._excel_tasks, "prepare_plan", None)
            if not callable(prepare):
                prepare = getattr(self._excel_tasks, "refine_plan", None)
            if not callable(prepare):
                prepare = self._excel_tasks.apply_plan
            prepare(
                plan,
                resolutions,
                operation="payment_sync",
            )
            return

        new_rows = list(_attribute(plan, "new_rows", default=()) or ())
        if new_rows and "selected_new_rows" not in resolutions:
            new_dialog = PaymentNewRowsDialog(new_rows, self)
            if new_dialog.exec() != QDialog.DialogCode.Accepted:
                self._excel_tasks.cancel_waiting()
                return
            resolutions["selected_new_rows"] = new_dialog.selected_item_ids
            if getattr(self, "_excel_review_plan", None) is not None:
                review_resolutions = getattr(
                    self, "_excel_review_resolutions", {}
                )
                review_resolutions["selected_new_rows"] = list(
                    resolutions["selected_new_rows"]
                )
                setattr(self, "_excel_review_resolutions", review_resolutions)
            MainWindow._save_excel_draft(
                self, plan, resolutions, "payment_sync"
            )

        if not bool(getattr(self, "_excel_review_confirmed", False)):
            summary = payment_confirmation_summary(
                plan,
                selected_new_ids=resolutions.get("selected_new_rows"),
            )
            dialog = ExcelConfirmationDialog(
                summary,
                confirm_label="Đồng bộ sang Thanh toán",
                parent=self,
            )
            if dialog.exec() != QDialog.DialogCode.Accepted:
                self._excel_tasks.cancel_waiting()
                return
            if getattr(self, "_excel_review_plan", None) is not None:
                setattr(self, "_excel_review_confirmed", True)
        self._excel_tasks.apply_plan(
            plan,
            resolutions,
            operation="payment_sync",
        )

    @Slot(object)
    def _excel_completed(self, result: Any) -> None:
        MainWindow._clear_excel_review(self)
        if self._excel_context == "configuration":
            valid = bool(_attribute(result, "is_valid", default=False))
            checks = list(_attribute(result, "checks", default=()) or ())
            detail = "\n".join(
                (
                    "✓ " if bool(_attribute(check, "ok", default=False)) else "✗ "
                )
                + str(_attribute(check, "message", default=""))
                for check in checks
            )
            self.settings_page.show_check_result(
                valid,
                detail or ("Cấu hình Excel hợp lệ." if valid else "Cấu hình Excel chưa hợp lệ."),
            )
            return

        operation = self._excel_operation or "sync"
        try:
            operation = self._excel_tasks.normalize_operation(
                _attribute(result, "operation", default=operation)
            )
        except Exception:
            pass
        self.workflow_page.set_excel_result(operation, result)
        self._load_excel_history()
        item_outcomes = list(_attribute(result, "item_outcomes", default=()) or ())
        latest = getattr(self, "_latest_excel_outcomes", None)
        if not isinstance(latest, dict):
            latest = {}
            self._latest_excel_outcomes = latest
        latest[operation] = item_outcomes
        availability = getattr(
            self.workflow_page, "set_latest_excel_data_available", None
        )
        if callable(availability):
            availability(operation, bool(item_outcomes))
        if operation == "payment_sync":
            summary = payment_completion_summary(result)
            dialog = ExcelCompletionDialog(
                summary,
                open_label="Mở file Thanh toán",
                details_available=bool(item_outcomes),
                parent=self,
            )
            dialog.exec()
            if dialog.selected_action == ExcelCompletionDialog.OPEN_TARGET:
                self._open_workbook_path(
                    _attribute(result, "target_path", default=None),
                    label="file Thanh toán",
                )
            elif dialog.selected_action == ExcelCompletionDialog.VIEW_DETAILS:
                ExcelOutcomeDialog(item_outcomes, parent=self).exec()
            return
        if operation == "sync":
            summary = daily_completion_summary(result)
        else:
            summary = posting_completion_summary(result)
        dialog = ExcelCompletionDialog(
            summary,
            open_label="Mở file BK",
            details_available=bool(item_outcomes),
            parent=self,
        )
        dialog.exec()
        if dialog.selected_action == ExcelCompletionDialog.OPEN_TARGET:
            self._open_bk_workbook(_attribute(result, "target_path", default=None))
        elif dialog.selected_action == ExcelCompletionDialog.VIEW_DETAILS:
            ExcelOutcomeDialog(item_outcomes, parent=self).exec()

    @Slot(object)
    def _excel_failed(self, error: Any) -> None:
        if self._excel_context == "configuration":
            self.settings_page.show_check_result(False, str(error))
            return
        retry_context = MainWindow._excel_retry_context_for(self, error)
        # A source reload owns this flag only until its new analysis either
        # reaches the review screen or fails.  Do not leak the highlight mode
        # into a later, unrelated posting run.
        self._excel_reanalysis_from_source_change = False
        MainWindow._clear_excel_review(self)
        self._show_excel_error(
            error,
            operation=self._excel_operation,
            retry_context=retry_context,
        )

    def _excel_retry_context_for(
        self, error: Any
    ) -> _ExcelRetryContext | None:
        if not isinstance(error, SourceDataChangedError):
            return None
        plan = getattr(self, "_excel_review_plan", None)
        if plan is None:
            return None
        repost_done = bool(
            _attribute(plan, "repost_selection_done", default=False)
        )
        repost_indices = (
            tuple(
                sorted(
                    int(value)
                    for value in (
                        _attribute(
                            plan, "repost_source_indices", default=()
                        )
                        or ()
                    )
                )
            )
            if repost_done
            else None
        )
        raw_batch_id = _attribute(plan, "batch_id", default=None)
        return _ExcelRetryContext(
            operation="posting",
            batch_id=int(raw_batch_id) if raw_batch_id is not None else None,
            sheet_name=_attribute(
                plan, "selected_sheet", "selected_sheet_name", default=None
            ),
            repost_source_indices=repost_indices,
            group_target_sheets=tuple(
                sorted(
                    (
                        str(_attribute(group, "group_id")),
                        str(_attribute(group, "target_sheet")),
                    )
                    for group in (
                        _attribute(plan, "source_groups", default=()) or ()
                    )
                    if _attribute(group, "group_id", default=None) not in (None, "")
                    and _attribute(group, "target_sheet", default=None) not in (None, "")
                )
            )
            or None,
            split_document_ids=tuple(
                sorted(
                    str(value)
                    for value in (
                        _attribute(plan, "split_document_ids", default=()) or ()
                    )
                    if str(value).strip()
                )
            )
            or None,
            highlight_unresolved=True,
        )

    @Slot(object)
    def _excel_correction_required(self, outcome: Any) -> None:
        if self._excel_tasks is None:
            return
        dialog = self._excel_review_dialog
        if dialog is None or self._excel_review_plan is None:
            # Defensive fallback for callers that use the controller without the
            # normal initial review screen.
            dialog = ConflictResolutionDialog(
                _attribute(outcome, "conflicts", default=()) or (),
                self,
                operation=self._excel_operation,
                initial_resolutions=_attribute(
                    outcome, "resolutions", default={}
                )
                or {},
                issues=_attribute(outcome, "issues", default=()) or (),
            )
            self._excel_review_dialog = dialog
            self._excel_review_plan = _attribute(
                self._excel_tasks, "_base_plan", default=None
            )
            self._excel_review_operation = self._excel_operation

        conflicts = list(_attribute(outcome, "conflicts", default=()) or ())
        issues = list(_attribute(outcome, "issues", default=()) or ())
        returned_resolutions = dict(
            _attribute(outcome, "resolutions", default={}) or {}
        )
        self._excel_review_resolutions.update(returned_resolutions)
        dialog.set_review(
            conflicts,
            issues=issues,
            initial_resolutions=self._excel_review_resolutions,
        )

        affected = {
            str(conflict_id)
            for issue in issues
            for conflict_id in (
                _attribute(issue, "conflict_ids", default=()) or ()
            )
        }
        count = len(affected) or len(issues)
        message = QMessageBox(self)
        message.setIcon(QMessageBox.Icon.Warning)
        message.setWindowTitle("Chưa thể ghi dữ liệu vào file")
        message.setText(
            f"Có {count} dòng chưa hợp lệ.\n\n"
            "File gốc chưa bị thay đổi."
        )
        review_button = message.addButton(
            f"Quay lại sửa {count} dòng lỗi",
            QMessageBox.ButtonRole.AcceptRole,
        )
        message.addButton("Hủy", QMessageBox.ButtonRole.RejectRole)
        show = getattr(dialog, "show", None)
        if callable(show):
            show()
        raise_dialog = getattr(dialog, "raise_", None)
        if callable(raise_dialog):
            raise_dialog()
        message.exec()
        if message.clickedButton() is not review_button:
            self._excel_tasks.cancel_waiting()
            self._clear_excel_review()
            return

        hide = getattr(dialog, "hide", None)
        if callable(hide):
            hide()
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self._excel_tasks.cancel_waiting()
            self._clear_excel_review()
            return
        self._excel_review_resolutions.update(dialog.resolution_map())
        operation = self._excel_review_operation or self._excel_operation
        prepare = getattr(self._excel_tasks, "prepare_plan", None)
        if not callable(prepare):
            prepare = getattr(self._excel_tasks, "refine_plan", None)
        if not callable(prepare):
            prepare = self._excel_tasks.apply_plan
        prepare(
            self._excel_review_plan,
            self._excel_review_resolutions,
            operation=operation,
        )

    def _clear_excel_review(self) -> None:
        dialog = getattr(self, "_excel_review_dialog", None)
        setattr(self, "_excel_review_dialog", None)
        setattr(self, "_excel_review_plan", None)
        setattr(self, "_excel_review_resolutions", {})
        setattr(self, "_excel_review_operation", None)
        setattr(self, "_excel_review_confirmed", False)
        if dialog is not None:
            close = getattr(dialog, "close", None)
            if callable(close):
                close()
            delete_later = getattr(dialog, "deleteLater", None)
            if callable(delete_later):
                delete_later()

    @Slot(str)
    def _excel_finished(self, operation: str) -> None:
        if self._excel_context == "configuration":
            self.settings_page.set_checking(False)
        else:
            self.workflow_page.set_excel_idle(operation)
        self.statusBar().showMessage("Sẵn sàng", 3000)
        self._excel_operation = None
        self._excel_context = None

    def _show_excel_error(
        self,
        error: Any,
        *,
        operation: str | None = None,
        retry_context: _ExcelRetryContext | None = None,
    ) -> None:
        text = str(error)
        folded = text.casefold()
        if isinstance(error, SourceDataChangedError):
            title = "Dữ liệu JSON đã thay đổi"
            retry_label = "Đọc lại dữ liệu"
        elif isinstance(error, WorkbookChangedError):
            label = str(getattr(error, "label", "Workbook") or "Workbook")
            title = f"{label} đã thay đổi"
            retry_label = "Đọc lại"
        elif isinstance(error, WorkbookLockedError):
            title = "File Excel đang được sử dụng"
            retry_label = "Thử lại"
        elif "ready" in folded or "json" in folded and "không có" in folded:
            title = "Không có JSON đã xác nhận"
            retry_label = None
        elif "khóa" in folded or "đang được mở" in folded:
            title = "File BK đang được sử dụng"
            retry_label = "Thử lại"
        elif "thay đổi" in folded:
            title = "File BK đã thay đổi"
            retry_label = "Đọc lại"
        elif "hàng ngày" in folded and (
            "không tìm thấy" in folded or "không thể đọc" in folded
        ):
            title = "Không thể đọc file Hàng ngày"
            retry_label = "Thử lại"
        elif "quyền" in folded:
            title = "Không có quyền ghi BK"
            retry_label = None
        else:
            title = "Không thể xử lý workbook"
            retry_label = None
        LOGGER.error("%s: %s", title, text)
        message = QMessageBox(self)
        message.setIcon(QMessageBox.Icon.Critical)
        message.setWindowTitle(title)
        message.setText(f"{text}\n\nFile BK gốc không bị thay đổi.")
        retry_button = None
        if retry_label is not None:
            retry_button = message.addButton(
                retry_label, QMessageBox.ButtonRole.AcceptRole
            )
            message.addButton("Hủy", QMessageBox.ButtonRole.RejectRole)
        elif title == "Không có JSON đã xác nhận":
            review_button = message.addButton(
                "Quay lại Bước 2", QMessageBox.ButtonRole.AcceptRole
            )
            message.addButton("Đóng", QMessageBox.ButtonRole.RejectRole)
        else:
            review_button = None
            message.addButton("Đóng", QMessageBox.ButtonRole.RejectRole)
        message.exec()
        if retry_button is not None and message.clickedButton() is retry_button:
            try:
                retry_operation = (
                    self._excel_tasks.normalize_operation(operation)
                    if self._excel_tasks is not None
                    else operation
                )
            except Exception:
                retry_operation = operation
            if retry_operation == "sync":
                QTimer.singleShot(0, self.start_daily_sync)
            elif retry_operation == "posting":
                if retry_context is not None:
                    QTimer.singleShot(
                        0,
                        lambda context=retry_context: self._restart_expense_posting(
                            context
                        ),
                    )
                else:
                    QTimer.singleShot(0, self.start_expense_posting)
            elif retry_operation == "payment_sync":
                QTimer.singleShot(0, self.start_payment_sync)
        elif (
            title == "Không có JSON đã xác nhận"
            and review_button is not None
            and message.clickedButton() is review_button
        ):
            self.navigation.setCurrentRow(0)
            if self._active_batch is not None:
                self.open_review(self._active_batch)

    def _restart_expense_posting(self, context: _ExcelRetryContext) -> None:
        if self._excel_tasks is None:
            return
        self._excel_context = "workflow"
        self._excel_reanalysis_from_source_change = bool(
            context.highlight_unresolved
        )
        kwargs: dict[str, Any] = {}
        if context.batch_id is not None:
            kwargs["batch_id"] = context.batch_id
        if context.sheet_name:
            kwargs["sheet_name"] = context.sheet_name
        if context.group_target_sheets:
            kwargs["group_target_sheets"] = dict(context.group_target_sheets)
        if context.split_document_ids:
            kwargs["split_document_ids"] = context.split_document_ids
        if context.repost_source_indices is not None:
            kwargs["repost_source_indices"] = context.repost_source_indices
        try:
            self._excel_tasks.start_posting(**kwargs)
        except Exception as exc:
            self._excel_reanalysis_from_source_change = False
            self._show_excel_error(exc, operation="posting")

    def _load_excel_history(self) -> None:
        repository = self._excel_run_repository
        if repository is None:
            return
        for operation, ui_operation in (
            ("DAILY_SYNC", "sync"),
            ("EXPENSE_POSTING", "posting"),
            ("PAYMENT_SYNC", "payment_sync"),
        ):
            try:
                record = repository.get_latest(
                    operation=operation,
                    statuses=("SUCCEEDED", "NO_CHANGES"),
                )
            except Exception:
                LOGGER.exception("Không đọc được lịch sử %s.", operation)
                continue
            if record is None:
                self._latest_excel_outcomes.pop(ui_operation, None)
                availability = getattr(
                    self.workflow_page, "set_latest_excel_data_available", None
                )
                if callable(availability):
                    availability(ui_operation, False)
                continue
            outcomes = _attribute(record, "item_outcomes", default=()) or ()
            self._latest_excel_outcomes[ui_operation] = list(outcomes)
            availability = getattr(
                self.workflow_page, "set_latest_excel_data_available", None
            )
            if callable(availability):
                availability(ui_operation, bool(outcomes))
            timestamp = _attribute(record, "completed_at", "started_at", default="")
            try:
                from datetime import datetime

                parsed = datetime.fromisoformat(str(timestamp).replace("Z", "+00:00"))
                timestamp_text = parsed.strftime("%d/%m/%Y %H:%M")
            except (TypeError, ValueError):
                timestamp_text = str(timestamp)
            if ui_operation == "sync":
                summary = (
                    f"{timestamp_text} – thêm "
                    f"{_attribute(record, 'changed_items', default=0)} dòng vào "
                    f"{_attribute(record, 'sheet_name', default='—')}"
                )
            elif ui_operation == "posting":
                summary = (
                    f"{timestamp_text} – nhập "
                    f"{_attribute(record, 'changed_items', default=0)} khoản, bỏ qua "
                    f"{_attribute(record, 'skipped_items', default=0)}"
                )
            else:
                summary = (
                    f"{timestamp_text} – thay đổi "
                    f"{_attribute(record, 'changed_items', default=0)} dòng, "
                    f"bỏ qua {_attribute(record, 'skipped_items', default=0)}"
                )
            self.workflow_page.set_excel_result(ui_operation, summary)

    @Slot(str)
    def show_latest_excel_data(self, operation: str) -> None:
        try:
            normalized = self.workflow_page._excel_operation(operation)
        except Exception:
            normalized = str(operation)
        outcomes = list(self._latest_excel_outcomes.get(normalized, ()))
        if not outcomes:
            QMessageBox.information(
                self,
                "Chưa có dữ liệu",
                "Chưa có dữ liệu hoàn tất nào của luồng này để xem lại.",
            )
            return
        ExcelOutcomeDialog(outcomes, parent=self).exec()

    def _load_rpa_history(self) -> None:
        payload: dict[str, Any] | None = None
        loader = getattr(self._rpa_expense, "load_latest_launched", None)
        if callable(loader):
            try:
                value = loader()
                payload = dict(value) if isinstance(value, Mapping) else None
            except Exception:
                LOGGER.exception("Không đọc được dữ liệu RPA gần nhất.")
        self._latest_rpa_payload = payload
        availability = getattr(
            self.workflow_page, "set_latest_rpa_data_available", None
        )
        if callable(availability):
            availability(bool(payload and payload.get("items")))

    @Slot()
    def show_latest_rpa_data(self) -> None:
        payload = self._latest_rpa_payload
        if not payload:
            QMessageBox.information(
                self,
                "Chưa có dữ liệu RPA",
                "Chưa có dữ liệu nào đã được gửi sang PAD để xem lại.",
            )
            return
        RpaLatestDataDialog(payload, self).exec()

    @Slot(object)
    def check_settings(self, settings_data: Mapping[str, Any]) -> None:
        launcher = self._assistant_launcher
        if launcher is None:
            self.settings_page.show_check_result(
                False, "Dịch vụ mở Trợ lý ảo chưa được khởi tạo."
            )
            return
        try:
            launcher.validate_configuration(
                settings_data.get("assistant_bat_path"),
                settings_data.get("output_dir"),
            )
            bang_ke_bat = str(
                settings_data.get("bang_ke_assistant_bat_path") or ""
            ).strip()
            if bang_ke_bat:
                launcher.validate_configuration(
                    bang_ke_bat,
                    settings_data.get("output_dir"),
                )
            rpa_expense_bat = str(
                settings_data.get("rpa_expense_bat_path") or ""
            ).strip()
            if rpa_expense_bat:
                from app.rpa_expense import RpaExpenseBatLauncher

                candidate_settings = self._build_settings(settings_data)
                RpaExpenseBatLauncher(
                    candidate_settings
                ).validate_configuration()
        except Exception as exc:
            self.settings_page.show_check_result(False, str(exc))
            return

        daily = str(settings_data.get("daily_workbook_path") or "").strip()
        bk = str(settings_data.get("bk_workbook_path") or "").strip()
        payment = str(
            settings_data.get("payment_workbook_path") or ""
        ).strip()
        if not daily or not bk or not payment:
            self.settings_page.show_check_result(
                False,
                "Cấu hình Trợ lý hợp lệ, nhưng chưa chọn đủ file Hàng ngày và file BK.",
            )
            return
        if self._excel_tasks is None:
            self.settings_page.show_check_result(
                False, "Dịch vụ kiểm tra workbook chưa được khởi tạo."
            )
            return
        try:
            from app.services.excel import ExcelConfigurationService

            candidate_settings = self._build_settings(settings_data)
            service = ExcelConfigurationService(candidate_settings)
            self._excel_context = "configuration"
            self.settings_page.set_checking(True)
            self._excel_tasks.submit(
                "sync",
                service.validate,
                with_progress=True,
            )
        except Exception as exc:
            self._excel_context = None
            self.settings_page.set_checking(False)
            self.settings_page.show_check_result(False, str(exc))

    def _launch_assistant(
        self,
        settings: Any,
        *,
        context: str,
        reconciliation_group_id: int | None = None,
        bat_setting: str = "assistant_bat_path",
    ) -> None:
        launcher = self._assistant_launcher
        if launcher is None:
            QMessageBox.warning(
                self,
                "Không có dịch vụ mở Trợ lý ảo",
                "Ứng dụng chưa khởi tạo được dịch vụ mở Trợ lý ảo.",
            )
            return
        active_session = self._next_assistant_session()
        if active_session is not None:
            QMessageBox.information(
                self,
                "Trợ lý ảo đang mở",
                "Đang có một cửa sổ Trợ lý chờ nhận file. "
                "Hãy hoàn tất hoặc đóng cửa sổ đó trước khi mở phiên mới.",
            )
            return
        selected_bat = str(
            _attribute(settings, bat_setting, default="") or ""
        ).strip()
        if not selected_bat and context == "bang_ke":
            QMessageBox.warning(
                self,
                "Chưa cấu hình BAT",
                "Chưa cấu hình file BAT mở tool bảng kê. "
                "Hãy chọn file trong Cài đặt.",
            )
            return
        try:
            result = launcher.launch(
                bat_path=selected_bat,
                output_dir=_attribute(settings, "output_dir", default=""),
                context=context,
                reconciliation_group_id=reconciliation_group_id,
            )
            session_id = str(_attribute(result, "session_id", default="") or "")
            if session_id:
                self._assistant_sessions.append(
                    _AssistantSession(
                        session_id=session_id,
                        context=context,
                        reconciliation_group_id=reconciliation_group_id,
                    )
                )
            self.statusBar().showMessage(
                str(_attribute(result, "message", default="Đã mở Trợ lý ảo.")),
                7000,
            )
        except Exception as exc:
            LOGGER.exception("Không mở được Trợ lý ảo: %s", exc)
            QMessageBox.warning(
                self,
                "Không mở được Trợ lý ảo",
                f"{exc}\nHãy kiểm tra file BAT và thư mục Output trong Cài đặt.",
            )

    @Slot(object)
    def _open_directory(self, path: Any) -> None:
        if not path:
            QMessageBox.information(
                self,
                "Chưa cấu hình thư mục",
                "Hãy chọn thư mục Output trong trang Cài đặt.",
            )
            return
        directory = Path(path).expanduser()
        if not directory.exists() or not directory.is_dir():
            QMessageBox.warning(
                self,
                "Không tìm thấy thư mục",
                f"Thư mục chưa tồn tại hoặc không truy cập được:\n{directory}",
            )
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(directory))):
            QMessageBox.warning(self, "Không mở được thư mục", str(directory))

    @Slot(object)
    def _open_bk_workbook(self, path: Any = None) -> None:
        raw_path = path or _attribute(
            self._settings, "bk_workbook_path", default=""
        )
        if not raw_path:
            QMessageBox.information(
                self,
                "Chưa cấu hình file BK",
                "Hãy chọn file BK trong trang Cài đặt.",
            )
            return
        workbook = Path(raw_path).expanduser()
        if not workbook.exists() or not workbook.is_file():
            QMessageBox.warning(
                self,
                "Không tìm thấy file BK",
                f"File chưa tồn tại hoặc không truy cập được:\n{workbook}",
            )
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(workbook))):
            QMessageBox.warning(self, "Không mở được file BK", str(workbook))

    def _open_workbook_path(self, path: Any, *, label: str) -> None:
        if not path:
            QMessageBox.information(
                self,
                f"Chưa cấu hình {label}",
                f"Hãy chọn {label} trong trang Cài đặt.",
            )
            return
        workbook = Path(path).expanduser()
        if not workbook.is_file():
            QMessageBox.warning(
                self,
                f"Không tìm thấy {label}",
                str(workbook),
            )
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(workbook))):
            QMessageBox.warning(
                self,
                f"Không mở được {label}",
                str(workbook),
            )

    @Slot(object)
    def open_containing_folder(self, path: Any) -> None:
        candidate = Path(path).expanduser()
        directory = candidate if candidate.is_dir() else candidate.parent
        self._open_directory(directory)

    @Slot(str)
    def _file_ready(self, raw_path: str) -> None:
        self.receive_file(Path(raw_path), manual=False)

    @Slot(str, object)
    def _file_processed(self, _raw_path: str, result: Any) -> None:
        self._apply_receive_result(result, automatic=True)

    def receive_file(self, path: str | Path, *, manual: bool = False) -> Any:
        service = self._batch_service
        if service is None:
            QMessageBox.warning(
                self,
                "Chưa thể tiếp nhận file",
                "Dịch vụ quản lý batch chưa được khởi tạo.",
            )
            return None
        handler = None
        for name in ("receive_file", "ingest_file", "process_file", "import_file"):
            candidate = getattr(service, name, None)
            if callable(candidate):
                handler = candidate
                break
        if handler is None:
            QMessageBox.warning(
                self,
                "Chưa thể tiếp nhận file",
                "Dịch vụ batch không hỗ trợ nhận file JSON.",
            )
            return None
        self.statusBar().showMessage(f"Đang tiếp nhận {Path(path).name}…")
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            result = handler(Path(path))
            self._apply_receive_result(result, automatic=not manual)
            return result
        except Exception as exc:
            LOGGER.exception("Tiếp nhận file thất bại, file=%s: %s", path, exc)
            QMessageBox.critical(
                self,
                "Không tiếp nhận được JSON",
                f"Không thể xử lý {Path(path).name}: {exc}\n"
                "Hãy xem trang Nhật ký để biết thêm.",
            )
            return None
        finally:
            QApplication.restoreOverrideCursor()

    def _apply_receive_result(self, result: Any, *, automatic: bool) -> None:
        if result is None:
            return
        duplicate = bool(_attribute(result, "duplicate", default=False))
        message = str(_attribute(result, "message", default="") or "")
        review = _attribute(result, "review")
        batch = _attribute(result, "batch", "metadata", default=result)
        target = review or batch

        if review is None and _batch_id(batch) is not None:
            review = self._load_batch_review(_batch_id(batch), show_error=False)
            target = review or batch
        self._set_active_batch(target)
        self.refresh_history(silent=True)
        if duplicate and not automatic:
            QMessageBox.information(
                self,
                "File đã được tiếp nhận",
                message or "File này đã được tiếp nhận trước đó. Batch cũ đã được mở lại.",
            )
        else:
            self.statusBar().showMessage(
                (
                    "File hiện hành đã được khôi phục."
                    if duplicate
                    else (message or "Đã tiếp nhận và kiểm tra file JSON.")
                ),
                8000,
            )
        if automatic and not duplicate and review is not None:
            session = self._next_assistant_session()
            dialog = self._sea_freight_dialog
            if (
                session is not None
                and session.context == "reconciliation"
                and self._sea_freight_service is not None
            ):
                group_id = session.reconciliation_group_id
                if group_id is None:
                    return
                metadata = _attribute(review, "metadata", default=batch)
                document = _attribute(review, "document")
                rows = list(_attribute(document, "rows", "data", "d", default=[]) or [])
                import_result = self._sea_freight_service.import_supplement(
                    group_id,
                    rows,
                    source_batch_id=_batch_id(metadata),
                    source_sha256=str(_attribute(metadata, "sha256", default="") or ""),
                )
                if (
                    dialog is not None
                    and dialog.isVisible()
                    and dialog.group_id == group_id
                ):
                    dialog.supplement_received(import_result)
                self._refresh_reconciliation_groups()
                if int(_attribute(import_result, "added_count", default=0) or 0) > 0:
                    self._complete_assistant_session(session)
            else:
                self._open_new_download(review)
                if session is not None and session.context in {"home", "bang_ke"}:
                    self._complete_assistant_session(session)

    def _next_assistant_session(self) -> _AssistantSession | None:
        is_active = getattr(self._assistant_launcher, "is_session_active", None)
        while self._assistant_sessions:
            session = self._assistant_sessions[0]
            if not callable(is_active):
                return session
            try:
                if is_active(session.session_id):
                    return session
            except Exception:
                LOGGER.exception("Không thể kiểm tra phiên Trợ lý ảo.")
                return session
            self._assistant_sessions.pop(0)
        return None

    def _complete_assistant_session(self, session: _AssistantSession) -> None:
        complete = getattr(self._assistant_launcher, "complete_session", None)
        if callable(complete):
            try:
                complete(session.session_id)
            except Exception:
                LOGGER.exception("Không thể yêu cầu đóng cửa sổ Trợ lý ảo.")
                return
        try:
            self._assistant_sessions.remove(session)
        except ValueError:
            pass

    def _open_new_download(self, review: Any) -> None:
        """Bỏ cửa sổ cũ và đưa dữ liệu vừa tải lên màn hình kiểm tra."""

        for window in list(self._review_windows.values()):
            window.model.mark_clean()
            window.close()
        self._review_windows.clear()
        self.navigation.setCurrentRow(0)
        self.showNormal()
        self.raise_()
        self.activateWindow()
        self.open_review(review)

    @Slot(object)
    def open_review(self, batch: Any = None) -> ReviewWindow | None:
        target = batch or self._active_batch
        if target is None:
            QMessageBox.information(
                self,
                "Chưa có batch",
                "Hãy tải file JSON từ Trợ lý ảo trước.",
            )
            return None
        batch_identifier = _batch_id(target)
        if batch_identifier in self._review_windows:
            window = self._review_windows[batch_identifier]
            window.showNormal()
            window.raise_()
            window.activateWindow()
            return window
        review = target
        if _attribute(target, "document") is None:
            review = self._load_batch_review(batch_identifier)
            if review is None:
                return None
        if _status_code(review) == "INVALID":
            metadata = _attribute(review, "metadata", default=review)
            detail = _attribute(metadata, "last_error", default="")
            QMessageBox.warning(
                self,
                "Batch không hợp lệ",
                "Batch này có lỗi cấu trúc nghiêm trọng nên không thể mở trình sửa dòng."
                + (f"\n\n{detail}" if detail else ""),
            )
            return None
        window = ReviewWindow(
            review,
            parent=self,
            batch_service=self._batch_service,
            validator=self._validator,
            sea_freight_service=self._sea_freight_service,
            settings=self._settings,
        )
        key = batch_identifier if batch_identifier is not None else id(window)
        self._review_windows[key] = window
        window.batchUpdated.connect(self._review_batch_updated)
        window.reconciliationChanged.connect(self._refresh_reconciliation_groups)
        window.reconciliationOpenRequested.connect(self.open_reconciliation)
        window.closed.connect(lambda key=key: self._review_windows.pop(key, None))
        window.show()
        window.raise_()
        window.activateWindow()
        return window

    def _refresh_reconciliation_groups(self) -> None:
        for window in list(self._review_windows.values()):
            window._restore_reconciliation_presentations()

    @Slot(int)
    def open_reconciliation(self, group_id: int) -> None:
        if self._sea_freight_service is None:
            QMessageBox.warning(self, "Chưa thể mở", "Dịch vụ đối soát chưa được khởi tạo.")
            return
        dialog = self._sea_freight_dialog
        if dialog is not None and dialog.isVisible():
            dialog.showNormal()
            dialog.raise_()
            dialog.activateWindow()
            if dialog.group_id != group_id:
                QMessageBox.information(
                    dialog,
                    "Hồ sơ đang được mở",
                    "Hãy đóng hồ sơ hiện tại trước khi mở hồ sơ của dòng khác.",
                )
            return
        if dialog is None:
            dialog = SeaFreightReconciliationDialog(
                self._sea_freight_service,
                batch_service=self._batch_service,
                open_assistant=lambda: self.open_reconciliation_assistant(
                    self._sea_freight_dialog.group_id
                ),
                group_id=group_id,
                parent=self,
            )
            dialog.changed.connect(self._refresh_reconciliation_groups)
            dialog.confirmed.connect(self._open_reconciliation_result)
            self._sea_freight_dialog = dialog
        else:
            dialog.load_group(group_id)
        dialog.showNormal()
        dialog.raise_()
        dialog.activateWindow()

    @Slot(object)
    def _open_reconciliation_result(self, review: Any) -> None:
        if self._sea_freight_dialog is not None:
            self._sea_freight_dialog.hide()
        self.refresh_history(silent=True)
        self._refresh_reconciliation_groups()
        self.statusBar().showMessage(
            "Đã xác nhận hồ sơ đối soát. File bóc tách gốc vẫn là phiên làm việc hiện hành.",
            8000,
        )

    def _load_batch_review(self, batch_identifier: Any, *, show_error: bool = True) -> Any:
        if self._batch_service is None or batch_identifier is None:
            return None
        for name in ("load_batch", "open_batch", "get_batch_review"):
            method = getattr(self._batch_service, name, None)
            if not callable(method):
                continue
            try:
                return method(batch_identifier)
            except Exception as exc:
                LOGGER.exception("Không mở được batch %s: %s", batch_identifier, exc)
                if show_error:
                    QMessageBox.critical(
                        self,
                        "Không mở được batch",
                        f"Không thể nạp batch #{batch_identifier}: {exc}",
                    )
                return None
        return None

    @Slot(object)
    def reload_batch(self, batch: Any = None) -> None:
        target = batch or self._active_batch
        identifier = _batch_id(target)
        existing = self._review_windows.get(identifier)
        if existing is not None and existing.model.dirty:
            QMessageBox.warning(
                self,
                "Có thay đổi chưa lưu",
                "Cửa sổ review đang có thay đổi chưa lưu. Hãy lưu hoặc đóng cửa sổ "
                "trước khi tải lại bản làm việc.",
            )
            existing.raise_()
            return
        review = self._load_batch_review(identifier)
        if review is None:
            return
        self._set_active_batch(review)
        if existing is not None:
            existing.replace_review(review)
            existing.raise_()
        self.statusBar().showMessage("Đã tải lại dữ liệu từ bản làm việc.", 5000)

    @Slot(object)
    def _review_batch_updated(self, result: Any) -> None:
        if _batch_id(result) == _batch_id(self._active_batch):
            self._set_active_batch(result)
        self.refresh_history(silent=True)

    @Slot(object)
    def _external_batch_changed(self, batch: Any) -> None:
        if batch is not None:
            self._set_active_batch(batch)
            self.refresh_history(silent=True)

    def _set_active_batch(self, batch: Any | None) -> None:
        self._active_batch = batch
        self.workflow_page.set_active_batch(batch)
        if batch is not None:
            service = self._batch_service
            method = getattr(service, "set_active_batch", None) if service is not None else None
            if callable(method):
                try:
                    method(_batch_id(batch))
                except Exception:
                    LOGGER.exception("Không lưu được active batch.")
        self.activeBatchChanged.emit(batch)

    @Slot()
    def refresh_history(self, *, silent: bool = False) -> None:
        if self._batch_service is None:
            self.history_page.set_batches([])
            return
        batches: Any = []
        for name in ("list_batches", "get_batches", "history"):
            method = getattr(self._batch_service, name, None)
            if not callable(method):
                continue
            try:
                batches = method()
            except TypeError:
                batches = method(limit=None)
            except Exception as exc:
                LOGGER.exception("Không tải được lịch sử: %s", exc)
                if not silent:
                    QMessageBox.warning(
                        self,
                        "Không tải được lịch sử",
                        "Hãy kiểm tra kết nối cơ sở dữ liệu và thử lại.",
                    )
                return
            break
        visible = [
            batch
            for batch in (batches or [])
            if str(_attribute(batch, "source_kind", default="ASSISTANT"))
            != "SEA_FREIGHT_RECONCILIATION"
        ]
        self.history_page.set_batches(visible)

    @Slot(object)
    def save_settings(self, data: Mapping[str, Any]) -> None:
        try:
            old_output = str(_attribute(self._settings, "output_dir", default=""))
            new_settings = self._build_settings(data)
            result = self._persist_settings(new_settings)
            if result is not None and not isinstance(result, (str, Path, bool)):
                new_settings = result
            self._settings = new_settings
            self._paths = _attribute(new_settings, "paths", default=self._paths)
            self.settings_page.mark_saved(new_settings)
            self.workflow_page.set_configuration(new_settings)
            update_launcher = getattr(self._assistant_launcher, "update_settings", None)
            if callable(update_launcher):
                update_launcher(new_settings)
            if self._batch_service is not None and hasattr(
                self._batch_service, "max_file_size_bytes"
            ):
                max_bytes = _attribute(new_settings, "max_file_size_bytes")
                if max_bytes is None:
                    max_megabytes = int(
                        _attribute(new_settings, "max_file_size_mb", default=50)
                    )
                    max_bytes = max_megabytes * 1024 * 1024
                self._batch_service.max_file_size_bytes = int(max_bytes)

            new_output = str(_attribute(new_settings, "output_dir", default=""))
            if self._watcher is not None and old_output != new_output:
                update_watcher = getattr(self._watcher, "update_settings", None) or getattr(
                    self._watcher, "configure", None
                )
                if callable(update_watcher):
                    try:
                        update_watcher(new_settings, restart=True)
                    except TypeError:
                        update_watcher(new_settings)
                else:
                    restart = getattr(self._watcher, "restart", None)
                    if callable(restart):
                        restart(_attribute(new_settings, "output_dir", default=""))
                current_output = getattr(
                    self._batch_service, "get_current_output_batch", None
                )
                current_batch = (
                    current_output() if callable(current_output) else None
                )
                self._set_active_batch(current_batch)
            self.settingsChanged.emit(new_settings)
            self.statusBar().showMessage("Đã lưu cấu hình.", 6000)
        except Exception as exc:
            LOGGER.exception("Không lưu được cấu hình: %s", exc)
            QMessageBox.critical(
                self,
                "Không lưu được cấu hình",
                f"{exc}\nHãy kiểm tra các đường dẫn và quyền ghi rồi thử lại.",
            )

    def _build_settings(self, data: Mapping[str, Any]) -> Any:
        current = self._settings
        merged: dict[str, Any] = {}
        to_dict = getattr(current, "to_dict", None)
        if callable(to_dict):
            merged.update(to_dict())
        elif isinstance(current, Mapping):
            merged.update(current)
        merged.update(dict(data))
        try:
            from app.config import AppSettings

            fallback_paths = _attribute(current, "paths")
            return AppSettings.from_dict(merged, fallback_paths=fallback_paths)
        except ImportError:
            return merged

    def _persist_settings(self, settings: Any) -> Any:
        controller_method = None
        if self._controller is not None:
            for name in ("apply_settings", "update_settings", "save_settings"):
                candidate = getattr(self._controller, name, None)
                if callable(candidate):
                    controller_method = candidate
                    break
        if controller_method is not None:
            return controller_method(settings)
        owner = self._config_manager
        if owner is not None:
            for name in ("save", "save_settings", "update"):
                method = getattr(owner, name, None)
                if callable(method):
                    return method(settings)
        raise RuntimeError("Chưa kết nối dịch vụ lưu cấu hình.")

    def _start_watcher(self) -> None:
        watcher = self._watcher
        if watcher is None:
            self._watcher_status_changed(
                False, "Bộ theo dõi Output chưa được khởi tạo."
            )
            return
        is_running = bool(_attribute(watcher, "is_running", default=False))
        if is_running:
            output = _attribute(watcher, "output_dir", default="")
            self._watcher_status_changed(True, f"Đang theo dõi: {output}")
            return
        start = getattr(watcher, "start", None)
        if callable(start):
            try:
                started = start()
                if started is False and not bool(
                    _attribute(watcher, "is_running", default=False)
                ):
                    self._watcher_status_changed(
                        False, "Không thể khởi động bộ theo dõi Output."
                    )
            except Exception as exc:
                LOGGER.exception("Không khởi động được watcher: %s", exc)
                self._watcher_status_changed(False, "Không thể theo dõi Output.")

    @Slot(bool, str)
    def _watcher_status_changed(self, running: bool, message: str) -> None:
        color = "#BEE3CC" if running else "#F6C9C5"
        background = "#19452F" if running else "#572A32"
        self.watcher_status.setStyleSheet(
            f"color: {color}; background: {background}; border-radius: 6px; padding: 8px;"
        )
        self.watcher_status.setText(message)

    @Slot(str, str)
    def _watcher_rejected(self, path: str, message: str) -> None:
        QMessageBox.warning(
            self,
            "Không thể tiếp nhận file",
            f"{Path(path).name}: {message}\nHãy kiểm tra file vừa tải xuống.",
        )

    @Slot(str)
    def _watcher_error(self, message: str) -> None:
        LOGGER.error("Watcher: %s", message)
        self.statusBar().showMessage(message, 10000)
        self._watcher_status_changed(False, message)

    @Slot(int)
    def _scan_completed(self, count: int) -> None:
        recorder = (
            getattr(self._controller, "record_output_scan", None)
            if self._controller is not None
            else None
        )
        if callable(recorder):
            try:
                recorder(count)
            except Exception:
                LOGGER.exception("Không lưu được thời điểm quét Output.")
        if count:
            self.statusBar().showMessage(
                f"Đã tìm thấy {count} file phù hợp trong Output; đang kiểm tra độ ổn định.",
                7000,
            )

    def _save_ui_state(self) -> None:
        owner = self._controller
        if owner is None:
            return
        payload = {
            "geometry": bytes(self.saveGeometry().toBase64()).decode("ascii"),
            "page": self.pages.currentIndex(),
        }
        for name in ("save_ui_state", "store_ui_state", "set_ui_state"):
            method = getattr(owner, name, None)
            if callable(method):
                try:
                    method(payload)
                except Exception:
                    LOGGER.exception("Không lưu được trạng thái cửa sổ.")
                return

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802
        if self._closing:
            event.accept()
            return
        for window in list(self._review_windows.values()):
            window.close()
            if window.isVisible():
                event.ignore()
                return
        self._closing = True
        self._save_ui_state()
        watcher = self._watcher
        stop = getattr(watcher, "stop", None) if watcher is not None else None
        if callable(stop):
            try:
                stop()
            except Exception:
                LOGGER.exception("Lỗi khi dừng watcher.")
        self.closing.emit()
        event.accept()
