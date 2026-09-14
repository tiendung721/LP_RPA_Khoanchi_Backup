"""Lớp nền dùng chung cho các cửa sổ con của ứng dụng."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QWidget


class AppDialog(QDialog):
    """QDialog có đầy đủ nút thu nhỏ và phóng to trên thanh tiêu đề."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowMinimizeButtonHint
            | Qt.WindowType.WindowMaximizeButtonHint
        )


__all__ = ["AppDialog"]
