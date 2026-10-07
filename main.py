"""Điểm khởi động ứng dụng Trợ lý Dữ liệu Quyết toán."""

from __future__ import annotations

import argparse
import logging
import sys
from collections.abc import Sequence
from pathlib import Path

from PySide6.QtCore import QLockFile
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QMessageBox

from app.application import ApplicationRuntime, configured_data_root
from app.config import AppPaths, software_root
from app.constants import APP_NAME
from app.logging_setup import install_exception_hook
from app.ui.main_window import MainWindow
from app.ui.theme import apply_application_theme

LOGGER = logging.getLogger(__name__)
APP_VERSION = "1.0.0"


def _parse_arguments(arguments: Sequence[str]) -> tuple[argparse.Namespace, list[str]]:
    parser = argparse.ArgumentParser(
        prog=Path(arguments[0]).name if arguments else "main.py",
        description="Trợ lý kiểm tra và duyệt JSON quyết toán từ thư mục Output.",
    )
    parser.add_argument(
        "--data-root",
        help=(
            "Thư mục gốc chứa Config, Database, Logs và Output\\_system. "
            "Có thể dùng biến môi trường TRO_LY_DATA_ROOT thay thế."
        ),
    )
    return parser.parse_known_args(list(arguments[1:]))


def _show_unhandled_error(message: str) -> None:
    parent = QApplication.activeWindow()
    QMessageBox.critical(parent, "Ứng dụng gặp lỗi", message)


def run(arguments: Sequence[str] | None = None) -> int:
    raw_arguments = list(arguments if arguments is not None else sys.argv)
    if not raw_arguments:
        raw_arguments = ["main.py"]
    options, qt_arguments = _parse_arguments(raw_arguments)
    application = QApplication([raw_arguments[0], *qt_arguments])
    application.setApplicationName(APP_NAME)
    application.setApplicationDisplayName(APP_NAME)
    application.setOrganizationName("Kikai")
    application.setApplicationVersion(APP_VERSION)
    application.setQuitOnLastWindowClosed(True)
    icon = QIcon(str(software_root() / "Assets" / "KIKAI_PM.ico"))
    if not icon.isNull():
        application.setWindowIcon(icon)
    apply_application_theme(application)

    runtime: ApplicationRuntime | None = None
    lock: QLockFile | None = None
    try:
        data_root = configured_data_root(options.data_root)
        bootstrap_paths = (
            AppPaths.from_data_root(data_root)
            if data_root is not None
            else AppPaths.defaults()
        )
        bootstrap_paths.config_dir.mkdir(parents=True, exist_ok=True)
        lock_path = bootstrap_paths.config_dir / "application.lock"
        lock = QLockFile(str(lock_path))
        lock.setStaleLockTime(30_000)
        if not lock.tryLock(200):
            QMessageBox.information(
                None,
                "Ứng dụng đang chạy",
                "Đã có một cửa sổ Trợ lý Dữ liệu Quyết toán đang chạy "
                "trên cùng thư mục dữ liệu.",
            )
            return 2

        runtime = ApplicationRuntime(data_root)
        install_exception_hook(_show_unhandled_error)
        LOGGER.info("Khởi động %s %s", APP_NAME, APP_VERSION)

        window = MainWindow(controller=runtime)
        application.aboutToQuit.connect(runtime.close)
        window.show()
        exit_code = application.exec()
        LOGGER.info("Ứng dụng kết thúc với mã %s", exit_code)
        return int(exit_code)
    except Exception as exc:
        logging.getLogger(__name__).exception("Không thể khởi động ứng dụng")
        QMessageBox.critical(
            None,
            "Không thể khởi động",
            f"Ứng dụng không thể khởi tạo: {exc}\n"
            "Hãy kiểm tra quyền ghi thư mục dữ liệu và xem file log.",
        )
        if runtime is not None:
            runtime.close()
        return 1
    finally:
        if lock is not None and lock.isLocked():
            lock.unlock()


if __name__ == "__main__":
    raise SystemExit(run())
