from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

from PySide6.QtWidgets import QAbstractButton, QLabel, QLineEdit, QMessageBox

from app.application import ApplicationRuntime, configured_data_root
from app.config import AppPaths, AppSettings, ConfigManager
from app.constants import APP_STATE_LAST_OUTPUT_SCAN
from app.models import DataRow
from app.sea_freight import BkContainerSnapshot, ContainerRecord, iso6346_check_digit
from app.ui.main_window import MainWindow


def _container(serial: int) -> str:
    body = f"TSTU{serial:06d}"
    return body + str(iso6346_check_digit(body))


class _SeaMatcher:
    def __init__(self, path: Path, count: int) -> None:
        self.path = path
        self.count = count

    def sheet_names(self, _path: object) -> tuple[str, ...]:
        return ("T07 26",)

    def snapshot(self, *_args: object, **_kwargs: object) -> BkContainerSnapshot:
        return BkContainerSnapshot(
            bk_path=str(self.path.resolve()), bk_sheet="T07 26",
            vessel_voyage_raw="PROSPER 2625S", vessel_key="PROSPER",
            voyage_key="2625S", combined_key="PROSPER2625S",
            workbook_fingerprint="fp", snapshot_hash="snapshot",
            containers=tuple(
                ContainerRecord(_container(index), "T07 26", index + 10, index, None)
                for index in range(1, self.count + 1)
            ),
        )


class _AssistantLauncher:
    def __init__(self) -> None:
        self.launches: list[dict[str, object]] = []
        self.completed: list[str] = []

    def launch(self, **kwargs: object) -> SimpleNamespace:
        self.launches.append(dict(kwargs))
        return SimpleNamespace(
            message="Đã mở Trợ lý ảo.",
            session_id=f"test-session-{len(self.launches):020d}",
        )

    def is_session_active(self, _session_id: str) -> bool:
        return True

    def complete_session(self, session_id: str) -> bool:
        self.completed.append(session_id)
        return True


def _isolated_runtime(tmp_path: Path) -> ApplicationRuntime:
    data_root = tmp_path / "runtime"
    output_dir = tmp_path / "Output"
    paths = AppPaths.from_data_root(data_root, output_dir)
    ConfigManager(paths=paths).save(
        AppSettings(data_root=data_root, output_dir=output_dir)
    )
    return ApplicationRuntime(data_root)


def test_runtime_and_main_window_start_with_isolated_data_root(
    qtbot, tmp_path: Path
) -> None:
    runtime = _isolated_runtime(tmp_path)
    window = MainWindow(controller=runtime, start_watcher=False)
    qtbot.addWidget(window)

    try:
        window.show()
        qtbot.wait(20)

        assert window.isVisible()
        assert runtime.paths.settings_path.is_file()
        assert runtime.paths.database_path.is_file()
        assert runtime.paths.logs_dir.is_dir()
        assert runtime.paths.excel_temp_dir.is_dir()
        assert runtime.paths.excel_backup_dir.is_dir()
        assert runtime.paths.excel_reports_dir.is_dir()
        assert runtime.paths.rpa_dir.is_dir()
        assert runtime.daily_sync_service is not None
        assert runtime.expense_posting_service is not None
        assert runtime.payment_sync_service is not None
        assert runtime.excel_task_controller is not None
        assert runtime.rpa_expense_controller is not None
        assert window.pages.count() == 4
        assert window.workflow_page.open_assistant_button.text() == "Mở Trợ lý ảo"
        assert not hasattr(window.workflow_page, "open_inbox_button")
        assert not hasattr(window.workflow_page, "choose_file_button")
        assert not hasattr(window.workflow_page, "pending_groups_button")
        assert (
            window.workflow_page.sync_daily_button.text()
            == "Đồng bộ dữ liệu Hàng ngày"
        )
        assert (
            window.workflow_page.post_expenses_button.text()
            == "Nhập khoản chi vào BK"
        )
        assert (
            window.workflow_page.sync_payment_button.text()
            == "Đồng bộ BK → Thanh toán"
        )
        settings_edits = window.settings_page.findChildren(QLineEdit)
        assert {
            "assistantBatEdit",
            "bangKeAssistantBatEdit",
            "outputDirEdit",
            "rpaExpenseBatEdit",
        }.issubset({edit.objectName() for edit in settings_edits})
        assert {
            edit.objectName()
            for edit in settings_edits
            if not edit.objectName().startswith("qt_")
        } == {
            "assistantBatEdit",
            "bangKeAssistantBatEdit",
            "outputDirEdit",
            "dailyWorkbookEdit",
            "bkWorkbookEdit",
            "paymentWorkbookEdit",
            "rpaExpenseBatEdit",
        }
        visible_text = " ".join(
            widget.text()
            for widget_type in (QLabel, QAbstractButton)
            for widget in window.findChildren(widget_type)
        ).casefold()
        assert "inbox" not in visible_text
        assert "outlook" not in visible_text
        assert "power automate" not in visible_text
        assert "webview2" not in visible_text
        runtime.record_output_scan(0)
        assert runtime.repository.get_app_state(APP_STATE_LAST_OUTPUT_SCAN)
    finally:
        window.close()
        runtime.close()


def test_configured_data_root_prefers_cli_then_environment(tmp_path: Path) -> None:
    cli = tmp_path / "cli"
    env = tmp_path / "env"

    assert configured_data_root(cli, environment={"TRO_LY_DATA_ROOT": str(env)}) == cli
    assert configured_data_root(None, environment={"TRO_LY_DATA_ROOT": str(env)}) == env
    assert configured_data_root(None, environment={}) is None


def test_legacy_container_json_is_rejected_by_normal_invoice_watcher(
    tmp_path: Path,
) -> None:
    runtime = _isolated_runtime(tmp_path)
    result = runtime.paths.output_dir / "ket_qua_boc_tach_result.json"
    result.write_text(
        json.dumps({"containers": ["VSGU2250713"]}),
        encoding="utf-8",
    )

    try:
        received = runtime._receive_watcher_file(result)
        assert received is not None
        assert received.review is None
        assert received.batch.status.value == "INVALID"
    finally:
        runtime.close()


def test_new_download_automatically_opens_review_window(
    qtbot, tmp_path: Path
) -> None:
    runtime = _isolated_runtime(tmp_path)
    window = MainWindow(controller=runtime, start_watcher=False)
    assistant = _AssistantLauncher()
    window._assistant_launcher = assistant
    qtbot.addWidget(window)
    source = runtime.paths.output_dir / "ket_qua_boc_tach.json"
    source.write_text(
        json.dumps(
            {
                "v": 1,
                "d": [["DRYU3026167", None, "VTN", "CV", None, None, 100]],
            }
        ),
        encoding="utf-8",
    )

    try:
        window.open_assistant()
        assert assistant.launches[-1]["context"] == "home"
        result = runtime.batch_service.receive_file(source)
        window._apply_receive_result(result, automatic=True)
        qtbot.wait(20)

        assert result.batch.id in window._review_windows
        old_review = window._review_windows[result.batch.id]
        assert old_review.isVisible()
        assert assistant.completed == ["test-session-00000000000000000001"]
        old_review.model.mark_dirty()

        incoming = runtime.paths.output_dir / "ket_qua_boc_tach (1).json"
        incoming.write_text(
            json.dumps(
                {
                    "v": 1,
                    "d": [["GAOU2112422", None, "VTN", "CV", None, None, 200]],
                }
            ),
            encoding="utf-8",
        )
        replacement = runtime.batch_service.receive_file(incoming)
        window._apply_receive_result(replacement, automatic=True)
        qtbot.wait(20)

        assert replacement.batch.id != result.batch.id
        assert result.batch.id not in window._review_windows
        assert replacement.batch.id in window._review_windows
        assert window._review_windows[replacement.batch.id].isVisible()
        assert (
            window._review_windows[replacement.batch.id].model.rows()[0].amount
            == 200
        )
    finally:
        window.close()
        runtime.close()


def test_bang_ke_download_automatically_closes_assistant(
    qtbot, tmp_path: Path
) -> None:
    runtime = _isolated_runtime(tmp_path)
    window = MainWindow(controller=runtime, start_watcher=False)
    assistant = _AssistantLauncher()
    window._assistant_launcher = assistant
    qtbot.addWidget(window)
    source = runtime.paths.output_dir / "ket_qua_bang_ke.json"
    source.write_text(
        json.dumps(
            {
                "v": 1,
                "d": [["DRYU3026167", None, "VTN", "CV", None, None, 100]],
            }
        ),
        encoding="utf-8",
    )

    try:
        window._launch_assistant(
            SimpleNamespace(
                bang_ke_assistant_bat_path="tool-bang-ke.bat",
                output_dir=runtime.paths.output_dir,
            ),
            context="bang_ke",
            bat_setting="bang_ke_assistant_bat_path",
        )
        result = runtime.batch_service.receive_file(source)
        window._apply_receive_result(result, automatic=True)

        assert assistant.launches[-1]["context"] == "bang_ke"
        assert assistant.completed == ["test-session-00000000000000000001"]
        assert window._assistant_sessions == []
    finally:
        window.close()
        runtime.close()


def test_new_download_is_added_to_open_reconciliation_without_closing_dialog(
    qtbot, tmp_path: Path
) -> None:
    runtime = _isolated_runtime(tmp_path)
    runtime.sea_freight_service.matcher = _SeaMatcher(tmp_path / "BK.xlsx", 3)
    window = MainWindow(controller=runtime, start_watcher=False)
    assistant = _AssistantLauncher()
    window._assistant_launcher = assistant
    qtbot.addWidget(window)
    initial_row = DataRow(
        cont=None, bl="BL-1", fee="CB", rule="HD", amount=20_000_000,
        invoice_no="INV-1", carrier="HÃNG A",
        vessel_voyage_raw="PROSPER 2625S", vessel_name="PROSPER", voyage_no="2625S",
        invoice_container_count=2, container_count_basis="EXPLICIT", invoice_date="2026-07-01",
    )
    first_path = runtime.paths.output_dir / "ket_qua_boc_tach.json"
    first_path.write_text(
        json.dumps({"v": 3, "d": [initial_row.to_object()]}, ensure_ascii=False),
        encoding="utf-8",
    )
    try:
        first = runtime.batch_service.receive_file(first_path)
        group = runtime.sea_freight_service.open_or_create(
            initial_row, bk_path=tmp_path / "BK.xlsx", month=7, year=2026,
            source_batch_id=first.batch.id, source_item_index=0,
            source_sha256=first.batch.sha256,
        )
        window.open_reconciliation(group.id)
        assert window._sea_freight_dialog is not None
        assert window._sea_freight_dialog.isVisible()
        window._sea_freight_dialog.assistant_button.click()
        assert assistant.launches[-1]["context"] == "reconciliation"
        assert assistant.launches[-1]["reconciliation_group_id"] == group.id

        unrelated = initial_row.copy_with(
            invoice_no="INV-KHAC", vessel_voyage_raw="OTHER 0001N",
            vessel_name="OTHER", voyage_no="0001N",
        )
        unrelated_path = runtime.paths.output_dir / "ket_qua_boc_tach (1).json"
        unrelated_path.write_text(
            json.dumps({"v": 3, "d": [unrelated.to_object()]}, ensure_ascii=False),
            encoding="utf-8",
        )
        unrelated_result = runtime.batch_service.receive_file(unrelated_path)
        window._apply_receive_result(unrelated_result, automatic=True)
        assert assistant.completed == []

        supplement = initial_row.copy_with(
            bl="BL-2", invoice_no="INV-2", amount=10_000_000,
            invoice_container_count=1,
        )
        incoming = runtime.paths.output_dir / "ket_qua_boc_tach (2).json"
        incoming.write_text(
            json.dumps({"v": 3, "d": [supplement.to_object()]}, ensure_ascii=False),
            encoding="utf-8",
        )
        received = runtime.batch_service.receive_file(incoming)
        window._apply_receive_result(received, automatic=True)
        qtbot.wait(20)

        contributions = runtime.sea_freight_repository.list_contributions(group.id)
        assert [item.invoice_no for item in contributions] == ["INV-1", "INV-2"]
        assert window._sea_freight_dialog.isVisible()
        assert received.batch.id not in window._review_windows
        assert window._sea_freight_dialog.status_label.text().startswith("Đã nhận thêm 1 HĐ")
        assert assistant.completed == ["test-session-00000000000000000001"]
    finally:
        if window._sea_freight_dialog is not None:
            window._sea_freight_dialog.close()
        window.close()
        runtime.close()


def test_visible_reconciliation_cannot_switch_to_another_row(
    qtbot, tmp_path: Path, monkeypatch
) -> None:
    runtime = _isolated_runtime(tmp_path)
    runtime.sea_freight_service.matcher = _SeaMatcher(tmp_path / "BK.xlsx", 2)
    window = MainWindow(controller=runtime, start_watcher=False)
    qtbot.addWidget(window)
    source = DataRow(
        cont=None, bl="BL-1", fee="CB", rule="HD", amount=20_000_000,
        invoice_no="INV-1", carrier="HÃNG A",
        vessel_voyage_raw="PROSPER 2625S", vessel_name="PROSPER", voyage_no="2625S",
        invoice_container_count=2, container_count_basis="EXPLICIT", invoice_date="2026-07-01",
    )
    first = runtime.sea_freight_service.open_or_create(
        source,
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=None,
        source_item_index=0,
        source_sha256="source-1",
    )
    second = runtime.sea_freight_repository.upsert_snapshot(
        BkContainerSnapshot(
            bk_path=str((tmp_path / "BK.xlsx").resolve()),
            bk_sheet="T07 26",
            vessel_voyage_raw="NEW VISION 2613S",
            vessel_key="NEWVISION",
            voyage_key="2613S",
            combined_key="NEWVISION2613S",
            workbook_fingerprint="fp-2",
            snapshot_hash="snapshot-2",
            containers=(),
        ),
        month=7,
        year=2026,
    )
    notices: list[str] = []
    monkeypatch.setattr(
        QMessageBox,
        "information",
        lambda _parent, _title, message, *_args, **_kwargs: notices.append(message),
    )

    try:
        window.open_reconciliation(first.id)
        dialog = window._sea_freight_dialog
        assert dialog is not None and dialog.isVisible()

        window.open_reconciliation(second.id)

        assert window._sea_freight_dialog is dialog
        assert dialog.group_id == first.id
        assert dialog.group_value.text().startswith("PROSPER 2625S")
        assert notices == [
            "Hãy đóng hồ sơ hiện tại trước khi mở hồ sơ của dòng khác."
        ]

        dialog.close()
        assert not dialog.isVisible()
        window.open_reconciliation(second.id)

        assert window._sea_freight_dialog is dialog
        assert dialog.group_id == second.id
        assert dialog.group_value.text().startswith("NEW VISION 2613S")
    finally:
        if window._sea_freight_dialog is not None:
            window._sea_freight_dialog.close()
        window.close()
        runtime.close()
