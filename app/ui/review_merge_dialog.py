"""Xác nhận nhóm khoản chi có thông tin mâu thuẫn trước khi gộp."""

from __future__ import annotations

from collections.abc import Sequence
from typing import Any

from PySide6.QtWidgets import (
    QComboBox,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QPlainTextEdit,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from .app_dialog import AppDialog
from .review_merge import FIELD_LABELS, choices_for_field
from .review_table_model import RULE_CATALOG, ReviewRow


def _display_value(field: str, value: Any) -> str:
    if field == "source_document":
        source_id, name = value
        return f"{name} [{source_id}]"
    if field == "rule":
        return f"{value} – {RULE_CATALOG.get(value, 'Không xác định')}"
    return str(value)


def _amount_text(value: int) -> str:
    return f"{value:,}".replace(",", ".")


class ReviewMergeDialog(AppDialog):
    """Buộc chọn rõ giá trị giữ lại khi một dòng tổng chỉ chứa được một giá trị."""

    def __init__(
        self,
        rows: Sequence[ReviewRow],
        source_positions: Sequence[int],
        conflicts: Sequence[str],
        *,
        exact_duplicate: bool = False,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Kiểm tra trước khi gộp khoản chi")
        self.setMinimumWidth(560)
        self._selectors: dict[str, QComboBox] = {}
        layout = QVBoxLayout(self)
        total = sum(row.amount for row in rows)
        title = QLabel(
            f"{len(rows)} dòng cùng {rows[0].fee} / {rows[0].cont} → "
            f"{_amount_text(total)} đ"
        )
        title.setStyleSheet("font-weight: 700;")
        layout.addWidget(title)

        details = QPlainTextEdit()
        details.setReadOnly(True)
        details.setMaximumHeight(120)
        details.setPlainText(
            "\n".join(
                f"Dòng {position + 1}: {_amount_text(row.amount)} đ | "
                f"HĐ {row.invoice_no or '—'} | ngày {row.invoice_date or '—'} | "
                f"{row.source_document_name}"
                for position, row in zip(source_positions, rows)
            )
        )
        layout.addWidget(details)
        notices: list[str] = []
        if exact_duplicate:
            notices.append("Có dòng trùng hoàn toàn; hãy xác nhận đây là các khoản tiền cần cộng.")
        if rows[0].fee == "CB":
            notices.append("Cước biển dùng tiền tổng hóa đơn; chỉ gộp nếu các số tiền thực sự cộng dồn. Quy tắc giữ HD.")
        if conflicts:
            notices.append("Chọn thông tin sẽ giữ trên dòng tổng. Giá trị khác sẽ không còn trong JSON sau khi lưu.")
        for notice in notices:
            label = QLabel(notice)
            label.setWordWrap(True)
            layout.addWidget(label)

        form_widget = QWidget()
        form = QFormLayout(form_widget)
        for field in conflicts:
            if field == "rule" and rows[0].fee == "CB":
                continue
            combo = QComboBox()
            combo.addItem("Chọn giá trị giữ lại…", None)
            for value in choices_for_field(rows, field):
                combo.addItem(_display_value(field, value), value)
            combo.currentIndexChanged.connect(self._update_accept_state)
            form.addRow(f"{FIELD_LABELS[field]}:", combo)
            self._selectors[field] = combo
        if self._selectors:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setWidget(form_widget)
            scroll.setMinimumHeight(min(330, 58 * len(self._selectors) + 24))
            layout.addWidget(scroll)

        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Gộp các dòng")
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Giữ nguyên")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self._update_accept_state()

    def _update_accept_state(self, *_args: Any) -> None:
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(
            all(combo.currentIndex() > 0 for combo in self._selectors.values())
        )

    def selected_values(self) -> dict[str, Any]:
        return {field: combo.currentData() for field, combo in self._selectors.items()}
