"""Delegate hiển thị một nút hành động trực tiếp trong bảng review."""

from __future__ import annotations

from PySide6.QtCore import QEvent, QModelIndex, QRect, Signal
from PySide6.QtGui import QMouseEvent, QPainter
from PySide6.QtWidgets import (
    QApplication,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionButton,
    QStyleOptionViewItem,
)

from app.ui.review_table_model import ReviewTableModel


class InlineActionDelegate(QStyledItemDelegate):
    clicked = Signal(object)

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex) -> None:
        if not bool(index.data(ReviewTableModel.ACTION_VISIBLE_ROLE)):
            empty = QStyleOptionViewItem(option)
            self.initStyleOption(empty, index)
            empty.text = ""
            super().paint(painter, empty, index)
            return
        base = QStyleOptionViewItem(option)
        self.initStyleOption(base, index)
        base.text = ""
        super().paint(painter, base, index)
        style = option.widget.style() if option.widget is not None else QApplication.style()
        style.drawControl(
            QStyle.ControlElement.CE_PushButton,
            self._button_option(option, index),
            painter,
        )

    def editorEvent(self, event, model, option, index) -> bool:  # noqa: N802
        if (
            event.type() == QEvent.Type.MouseButtonRelease
            and isinstance(event, QMouseEvent)
            and bool(index.data(ReviewTableModel.ACTION_VISIBLE_ROLE))
            and bool(index.data(ReviewTableModel.ACTION_ENABLED_ROLE))
            and self._button_rect(option.rect).contains(event.position().toPoint())
        ):
            self.clicked.emit(index)
            return True
        return False

    @classmethod
    def _button_option(cls, option: QStyleOptionViewItem, index: QModelIndex) -> QStyleOptionButton:
        button = QStyleOptionButton()
        button.rect = cls._button_rect(option.rect)
        button.text = str(index.data() or "")
        button.state = QStyle.StateFlag.State_Enabled
        if not bool(index.data(ReviewTableModel.ACTION_ENABLED_ROLE)):
            button.state &= ~QStyle.StateFlag.State_Enabled
        return button

    @staticmethod
    def _button_rect(rect: QRect) -> QRect:
        return rect.adjusted(4, 3, -4, -3)


__all__ = ["InlineActionDelegate"]
