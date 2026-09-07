from __future__ import annotations

from PySide6.QtWidgets import QLabel, QScrollArea

from app.ui.settings_page import SettingsPage
from app.ui.theme import APP_STYLESHEET
from app.ui.workflow_page import WorkflowPage


def test_excel_loading_feedback_tracks_running_state(qtbot) -> None:
    page = WorkflowPage()
    qtbot.addWidget(page)

    assert page.excel_loading_bar.isHidden()

    page.set_excel_running("sync", "Đang đọc workbook")

    assert not page.excel_loading_bar.isHidden()
    assert page.sync_daily_button.property("loading") is True
    assert page.post_expenses_button.property("loading") is False

    page.set_excel_idle("sync")

    assert page.excel_loading_bar.isHidden()
    assert page.sync_daily_button.property("loading") is False


def test_rpa_loading_feedback_tracks_running_state(qtbot) -> None:
    page = WorkflowPage()
    qtbot.addWidget(page)

    page.set_rpa_running("Đang tổng hợp dữ liệu")

    assert not page.rpa_loading_bar.isHidden()
    assert page.run_rpa_expense_button.property("loading") is True

    page.set_rpa_idle()

    assert page.rpa_loading_bar.isHidden()
    assert page.run_rpa_expense_button.property("loading") is False


def test_settings_checking_feedback_restores_button(qtbot) -> None:
    page = SettingsPage()
    qtbot.addWidget(page)
    page.save_button.setEnabled(True)

    page.set_checking(True)

    assert not page.check_loading_bar.isHidden()
    assert page.check_button.text() == "Đang kiểm tra…"
    assert page.check_button.property("loading") is True
    assert not page.check_button.isEnabled()

    page.set_checking(False)

    assert page.check_loading_bar.isHidden()
    assert page.check_button.text() == "Kiểm tra cấu hình"
    assert page.check_button.property("loading") is False
    assert page.check_button.isEnabled()


def test_settings_page_fits_minimum_application_height_without_scrollbar(
    qtbot,
) -> None:
    page = SettingsPage()
    page.setStyleSheet(APP_STYLESHEET)
    page.resize(980, 600)
    qtbot.addWidget(page)

    page.show()
    qtbot.wait(10)

    scroll = page.findChild(QScrollArea)
    assert scroll is not None
    assert not scroll.verticalScrollBar().isVisible()
    assert scroll.verticalScrollBar().maximum() == 0


def test_workflow_uses_named_groups_instead_of_numbered_steps(qtbot) -> None:
    page = WorkflowPage()
    qtbot.addWidget(page)

    visible_copy = " ".join(label.text() for label in page.findChildren(QLabel))

    assert "Bước 1" not in visible_copy
    assert "Bước 2" not in visible_copy
    assert "Bước 3" not in visible_copy
    assert "Bước 4" not in visible_copy
    assert "BÓC TÁCH CHỨNG TỪ" in visible_copy
    assert "KIỂM TRA DỮ LIỆU" in visible_copy
    assert "XỬ LÝ EXCEL" in visible_copy
    assert "TỰ ĐỘNG HÓA RPA" in visible_copy


def test_excel_actions_are_three_distinct_vertical_rows(qtbot) -> None:
    page = WorkflowPage()
    qtbot.addWidget(page)

    action_rows = (
        page.daily_sync_action,
        page.expense_posting_action,
        page.payment_sync_action,
    )
    buttons = (
        page.sync_daily_button,
        page.post_expenses_button,
        page.sync_payment_button,
    )

    assert all(
        button.parentWidget() is row
        for row, button in zip(action_rows, buttons)
    )
    indexes = [page.step3_card.layout().indexOf(row) for row in action_rows]
    assert indexes == sorted(indexes)
    assert len(set(indexes)) == 3
