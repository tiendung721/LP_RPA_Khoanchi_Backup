"""Khởi tạo và quản lý vòng đời các service của ứng dụng."""

from __future__ import annotations

import logging
import os
from collections.abc import Mapping
from pathlib import Path
from threading import RLock
from typing import Any

from app.config import (
    AppPaths,
    AppSettings,
    ConfigManager,
    migrate_legacy_runtime_layout,
)
from app.constants import (
    APP_STATE_LAST_OUTPUT_SCAN,
    APP_STATE_LAST_PAGE,
    APP_STATE_MAIN_WINDOW_GEOMETRY,
)
from app.database import Database
from app.logging_setup import setup_logging
from app.repositories.batch_repository import BatchRepository
from app.repositories.batch_repository import local_now_iso
from app.repositories.excel_draft_repository import ExcelDraftRepository
from app.repositories.excel_run_repository import ExcelRunRepository
from app.repositories.expense_posting_repository import ExpensePostingRepository
from app.services.batch_service import BatchService
from app.services.assistant_bat_launcher import AssistantBatLauncher
from app.sea_freight import SeaFreightReconciliationService, SeaFreightRepository
from app.rpa_expense import (
    RpaChoiceService,
    RpaExpenseBatLauncher,
    RpaExpenseService,
)
from app.services.excel import (
    DailySyncService,
    ExcelConfigurationService,
    ExpensePostingService,
    PaymentSyncService,
)
from app.services.excel.drafts import ExcelDraftService
from app.services.output_watcher import OutputWatcher
from app.services.json_codec import JsonCodec
from app.services.reviewed_batch_provider import ReviewedBatchProvider
from app.services.validation_service import ValidationService
from app.ui.excel_task_controller import ExcelTaskController
from app.ui.rpa_expense_controller import RpaExpenseController

LOGGER = logging.getLogger(__name__)


class ApplicationRuntime:
    """Composition root duy nhất cho UI, service và tài nguyên runtime."""

    def __init__(self, data_root: str | Path | None = None) -> None:
        bootstrap_paths = (
            AppPaths.from_data_root(Path(data_root).expanduser())
            if data_root is not None
            else AppPaths.defaults()
        )
        bootstrap_paths.ensure_directories()
        setup_logging(bootstrap_paths)

        self.config_manager = ConfigManager(paths=bootstrap_paths)
        self.settings = self.config_manager.load()
        self.paths = self.settings.paths
        moved_roots = migrate_legacy_runtime_layout(self.paths)
        self.paths.ensure_directories()
        self.log_path = setup_logging(self.paths)

        self.database = Database(self.paths.database_path)
        changed_rows = 0
        if self.config_manager.relocated_from is not None:
            changed_rows += self.database.rebase_paths(
                self.config_manager.relocated_from,
                self.paths.data_root,
            )
        for old_root, new_root in self.paths.legacy_runtime_rebases:
            changed_rows += self.database.rebase_paths(old_root, new_root)
        if moved_roots or changed_rows:
            LOGGER.info(
                "Đã gom dữ liệu vào %s; chuyển %s thư mục và cập nhật %s bản ghi",
                self.paths.system_dir,
                moved_roots,
                changed_rows,
            )
        self.repository = BatchRepository(self.database)
        self.validation_service = ValidationService()
        self.validator = self.validation_service
        self.json_codec = JsonCodec()
        self.batch_service = BatchService(
            self.paths,
            self.repository,
            codec=self.json_codec,
            validation_service=self.validation_service,
            max_file_size_bytes=self.settings.max_file_size_bytes,
        )
        self.reviewed_batch_provider = ReviewedBatchProvider(self.batch_service)
        self.excel_run_repository = ExcelRunRepository(self.database)
        self.excel_draft_repository = ExcelDraftRepository(self.database)
        self.excel_draft_service = ExcelDraftService(self.excel_draft_repository)
        self.expense_posting_repository = ExpensePostingRepository(self.database)
        self.sea_freight_repository = SeaFreightRepository(self.database)
        self.sea_freight_service = SeaFreightReconciliationService(
            self.sea_freight_repository
        )
        self._configure_excel_services(self.settings)
        removed_excel_artifacts = self.daily_sync_service.cleanup_stale_files()
        if removed_excel_artifacts:
            LOGGER.info(
                "Đã dọn %s file Excel tạm còn sót từ lần chạy trước.",
                removed_excel_artifacts,
            )
        self.excel_task_controller = ExcelTaskController(
            daily_sync_service=self.daily_sync_service,
            expense_posting_service=self.expense_posting_service,
            payment_sync_service=self.payment_sync_service,
            draft_service=self.excel_draft_service,
        )
        self.rpa_expense_service = RpaExpenseService(self.settings)
        self.rpa_choice_service = RpaChoiceService(self.excel_draft_repository)
        self.rpa_expense_launcher = RpaExpenseBatLauncher(self.settings)
        self.rpa_expense_controller = RpaExpenseController(
            self.rpa_expense_service,
            self.rpa_expense_launcher,
            self.rpa_choice_service,
        )
        self.assistant_launcher = AssistantBatLauncher(self.settings)
        self.launcher = self.assistant_launcher
        self.watcher = OutputWatcher(
            self.settings,
            on_file_ready=self._receive_watcher_file,
        )
        self.output_watcher = self.watcher
        self.batch_service.set_output_write_callback(self.watcher.mark_handled)
        self._close_lock = RLock()
        self._closed = False
        LOGGER.info(
            "Khởi tạo runtime thành công; data_root=%s, output=%s",
            self.paths.data_root,
            self.settings.output_dir,
        )

    def _receive_watcher_file(self, path: Path) -> object:
        """Nhận callback đã được watcher chuyển an toàn về Qt owner thread."""

        LOGGER.info("Watcher chuyển file ổn định sang BatchService: %s", path.name)
        return self.batch_service.receive_file(path)

    def _configure_excel_services(self, settings: AppSettings) -> None:
        """Tạo các service Bước 3 từ cùng settings/repository của runtime."""

        self.daily_sync_service = DailySyncService(
            settings,
            run_repository=self.excel_run_repository,
        )
        self.expense_posting_service = ExpensePostingService(
            self.reviewed_batch_provider,
            settings,
            run_repository=self.excel_run_repository,
            posting_repository=self.expense_posting_repository,
            sea_freight_repository=self.sea_freight_repository,
            sea_freight_service=self.sea_freight_service,
        )
        self.payment_sync_service = PaymentSyncService(
            settings,
            run_repository=self.excel_run_repository,
        )
        self.excel_configuration_service = ExcelConfigurationService(settings)

    def apply_settings(self, settings: AppSettings) -> AppSettings:
        """Lưu cấu hình và cập nhật các service không cần tái tạo database."""

        if not isinstance(settings, AppSettings):
            raise TypeError("Cấu hình mới phải là AppSettings.")
        old_root = self.paths.data_root.resolve()
        new_paths = settings.paths
        if new_paths.data_root.resolve() != old_root:
            raise ValueError(
                "Không thể đổi data_root khi ứng dụng đang chạy. "
                "Hãy đóng ứng dụng và khởi động lại với --data-root."
            )
        task_controller = getattr(self, "excel_task_controller", None)
        if task_controller is not None and task_controller.is_busy:
            raise ValueError("Không thể đổi cấu hình khi tác vụ Excel đang chạy.")
        rpa_controller = getattr(self, "rpa_expense_controller", None)
        if rpa_controller is not None and rpa_controller.is_busy:
            raise ValueError(
                "Không thể đổi cấu hình khi luồng RPA đang chuẩn bị dữ liệu."
            )
        new_paths.ensure_directories()
        self.config_manager.save(settings)
        self.settings = settings
        self.paths = new_paths
        self.batch_service.update_paths(new_paths)
        self.batch_service.max_file_size_bytes = settings.max_file_size_bytes
        self.assistant_launcher.update_settings(settings)
        self._configure_excel_services(settings)
        if task_controller is not None:
            task_controller.update_services(
                daily_sync_service=self.daily_sync_service,
                expense_posting_service=self.expense_posting_service,
                payment_sync_service=self.payment_sync_service,
            )
        self.rpa_expense_service = RpaExpenseService(settings)
        self.rpa_expense_launcher = RpaExpenseBatLauncher(settings)
        if rpa_controller is not None:
            rpa_controller.update_services(
                service=self.rpa_expense_service,
                launcher=self.rpa_expense_launcher,
            )
        LOGGER.info(
            "Đã áp dụng cấu hình; bat=%s, output=%s",
            settings.assistant_bat_path,
            settings.output_dir,
        )
        return settings

    update_settings = apply_settings
    save_settings = apply_settings

    def save_ui_state(self, state: Mapping[str, Any]) -> None:
        geometry = state.get("geometry")
        page = state.get("page", state.get("page_index"))
        self.repository.set_app_state(
            APP_STATE_MAIN_WINDOW_GEOMETRY,
            str(geometry) if geometry else None,
        )
        self.repository.set_app_state(
            APP_STATE_LAST_PAGE,
            str(page) if isinstance(page, int) else None,
        )

    def restore_ui_state(self) -> dict[str, Any]:
        state: dict[str, Any] = {}
        geometry = self.repository.get_app_state(APP_STATE_MAIN_WINDOW_GEOMETRY)
        page = self.repository.get_app_state(APP_STATE_LAST_PAGE)
        if geometry:
            state["geometry"] = geometry
        if page is not None:
            try:
                state["page"] = int(page)
            except ValueError:
                self.repository.set_app_state(APP_STATE_LAST_PAGE, None)
        return state

    load_ui_state = restore_ui_state

    def record_output_scan(self, _candidate_count: int = 0) -> None:
        """Lưu thời điểm quét Output gần nhất để chẩn đoán/khôi phục trạng thái."""

        self.repository.set_app_state(APP_STATE_LAST_OUTPUT_SCAN, local_now_iso())

    def close(self) -> None:
        with self._close_lock:
            if self._closed:
                return
            self._closed = True
        LOGGER.info("Đang dừng ứng dụng")
        try:
            self.watcher.stop()
        except Exception:
            LOGGER.exception("Không thể dừng watcher sạch")
        try:
            self.excel_task_controller.shutdown(wait=True)
        except Exception:
            LOGGER.exception("Không thể dừng worker Excel sạch")
        try:
            self.rpa_expense_controller.shutdown(wait=True)
        except Exception:
            LOGGER.exception("Không thể dừng worker RPA sạch")
        try:
            self.assistant_launcher.close()
        except Exception:
            LOGGER.exception("Không thể dừng kết nối Trợ lý ảo sạch")
        try:
            self.database.close()
        except Exception:
            LOGGER.exception("Không thể đóng database sạch")
        logging.shutdown()


def configured_data_root(
    command_line_value: str | Path | None = None,
    *,
    environment: Mapping[str, str] | None = None,
) -> Path | None:
    """Ưu tiên CLI, sau đó biến môi trường; trả ``None`` để dùng mặc định."""

    if command_line_value:
        return Path(command_line_value).expanduser()
    env = environment if environment is not None else os.environ
    value = env.get("TRO_LY_DATA_ROOT", "").strip()
    return Path(value).expanduser() if value else None
