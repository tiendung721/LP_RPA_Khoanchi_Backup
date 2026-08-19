from __future__ import annotations

from PySide6.QtWidgets import QScrollArea

from app.ui.main_window import MainWindow
from app.ui.workflow_page import WorkflowPage


def test_workflow_only_exposes_assistant_action_in_step_one(qtbot) -> None:
    page = WorkflowPage({"assistant_bat_path": ""})
    qtbot.addWidget(page)

    assert page.open_assistant_button.text() == "Mở Trợ lý ảo"
    assert page.open_bang_ke_assistant_button.text() == "Mở tool bảng kê"
    assert not hasattr(page, "open_inbox_button")
    assert not hasattr(page, "choose_file_button")
    assert not hasattr(page, "pending_card")
    assert not hasattr(page, "pending_groups_button")
    assert page.review_button.isEnabled() is False


def test_valid_batch_shows_saved_message_and_vietnamese_time(qtbot) -> None:
    page = WorkflowPage()
    qtbot.addWidget(page)

    page.set_active_batch(
        {
            "id": 3,
            "source_filename": "ket_qua_boc_tach.json",
            "status": "REVIEWING",
            "last_saved_at": "2026-07-27T22:18:00+07:00",
        }
    )

    assert page.file_status_badge.text() == "Đã có file"
    assert page.file_name_label.text() == "ket_qua_boc_tach.json"
    assert "Đã lưu dữ liệu bóc tách JSON" in page.file_note_label.text()
    assert page.saved_label.text() == (
        "Lưu thành công lần cuối: 22:18 ngày 27/07/2026"
    )
    assert page.review_button.isEnabled()


def test_invalid_batch_never_claims_successful_save(qtbot) -> None:
    page = WorkflowPage()
    qtbot.addWidget(page)

    page.set_active_batch(
        {
            "id": 4,
            "source_filename": "ket_qua_boc_tach.json",
            "status": "INVALID",
            "last_error": "Thiếu khóa d.",
        }
    )

    assert page.file_status_badge.text() == "File không hợp lệ"
    assert page.file_note_label.text() == "Thiếu khóa d."
    assert page.saved_label.text().endswith("—")
    assert not page.review_button.isEnabled()


def test_four_latest_data_buttons_are_independent_and_emit_their_flow(qtbot) -> None:
    page = WorkflowPage()
    qtbot.addWidget(page)
    excel_requests: list[str] = []
    rpa_requests: list[bool] = []
    page.view_latest_excel_requested.connect(excel_requests.append)
    page.view_latest_rpa_requested.connect(lambda: rpa_requests.append(True))

    assert not page.view_daily_sync_button.isEnabled()
    assert not page.view_expense_posting_button.isEnabled()
    assert not page.view_payment_sync_button.isEnabled()
    assert not page.view_rpa_expense_button.isEnabled()
    assert page.view_daily_sync_button.text() == "Chưa có dữ liệu"
    assert page.view_rpa_expense_button.text() == "Chưa có dữ liệu đã gửi"

    for operation, button in (
        ("sync", page.view_daily_sync_button),
        ("posting", page.view_expense_posting_button),
        ("payment_sync", page.view_payment_sync_button),
    ):
        page.set_latest_excel_data_available(operation, True)
        button.click()
    page.set_latest_rpa_data_available(True)
    page.view_rpa_expense_button.click()

    assert excel_requests == ["sync", "posting", "payment_sync"]
    assert rpa_requests == [True]
    assert page.view_daily_sync_button.text() == "Xem lại ›"
    assert page.view_rpa_expense_button.text() == "Xem dữ liệu đã gửi ›"


def test_four_workflow_groups_fit_minimum_window_without_scroll(qtbot) -> None:
    window = MainWindow(settings={}, start_watcher=False)
    qtbot.addWidget(window)
    window.resize(980, 650)
    window.show()
    qtbot.wait(20)

    page = window.workflow_page
    cards = (
        page.step1_card,
        page.step2_card,
        page.step3_card,
        page.step4_card,
    )

    assert page.findChildren(QScrollArea) == []
    assert all(card.isVisible() for card in cards)
    assert all(page.rect().contains(card.geometry()) for card in cards)
    assert len({card.geometry().left() for card in cards}) == 1
    assert [card.geometry().top() for card in cards] == sorted(
        card.geometry().top() for card in cards
    )
    assert all(
        following.geometry().top() - previous.geometry().bottom() <= 16
        for previous, following in zip(cards, cards[1:])
    )
    assert page.open_assistant_button.width() == page.review_button.width()
    assert page.review_button.width() == page.run_rpa_expense_button.width()
    assert all(
        row.rect().contains(button.geometry())
        for row, button in (
            (page.daily_sync_action, page.sync_daily_button),
            (page.expense_posting_action, page.post_expenses_button),
            (page.payment_sync_action, page.sync_payment_button),
        )
    )
    status_right_edges = {
        badge.mapTo(page, badge.rect().topRight()).x()
        for badge in (
            page.assistant_status,
            page.file_status_badge,
            page.rpa_configuration_status,
        )
    }
    assert len(status_right_edges) == 1
