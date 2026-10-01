"""Các thành phần trình bày dùng chung cho hộp thoại desktop."""

from __future__ import annotations

from PySide6.QtWidgets import QApplication, QFrame, QLabel, QVBoxLayout, QWidget


def dialog_intro(title: str, description: str) -> tuple[QFrame, QLabel]:
    """Tạo phần giới thiệu ngắn, không tham gia xử lý dữ liệu."""

    frame = QFrame()
    frame.setObjectName("dialogIntro")
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(15, 11, 15, 11)
    layout.setSpacing(3)
    title_label = QLabel(title)
    title_label.setObjectName("dialogIntroTitle")
    title_label.setWordWrap(True)
    description_label = QLabel(description)
    description_label.setObjectName("dialogIntroDescription")
    description_label.setWordWrap(True)
    layout.addWidget(title_label)
    layout.addWidget(description_label)
    return frame, title_label


def fit_window_to_screen(window: QWidget, width: int, height: int) -> None:
    """Giới hạn kích thước mở ban đầu theo vùng làm việc của màn hình."""

    screen = window.screen() or QApplication.primaryScreen()
    if screen is None:
        window.resize(width, height)
        return
    area = screen.availableGeometry()
    window.resize(
        min(width, max(1, area.width() - 32)),
        min(height, max(1, area.height() - 32)),
    )
