"""Điều phối tác vụ nền cho luồng 4 nhập khoản chi bằng PAD."""

from __future__ import annotations

from concurrent.futures import Future, ThreadPoolExecutor
from threading import RLock
from typing import Any, Iterable

from PySide6.QtCore import QObject, Signal

from app.rpa_expense import RpaExpenseBatLauncher, RpaExpenseService


class RpaExpenseController(QObject):
    started = Signal(str)
    progress = Signal(str)
    sheets_ready = Signal(object)
    plan_ready = Signal(object)
    launched = Signal(object)
    failed = Signal(object)
    finished = Signal(str)

    def __init__(
        self,
        service: RpaExpenseService,
        launcher: RpaExpenseBatLauncher,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self.service = service
        self.launcher = launcher
        self._executor = ThreadPoolExecutor(
            max_workers=1,
            thread_name_prefix="rpa-expense",
        )
        self._lock = RLock()
        self._phase: str | None = None

    @property
    def is_busy(self) -> bool:
        with self._lock:
            return self._phase is not None

    def update_services(
        self,
        *,
        service: RpaExpenseService,
        launcher: RpaExpenseBatLauncher,
    ) -> None:
        if self.is_busy:
            raise RuntimeError(
                "Không thể đổi cấu hình khi luồng RPA đang chuẩn bị dữ liệu."
            )
        self.service = service
        self.launcher = launcher

    def load_sheets(self) -> None:
        self._submit(
            "sheets",
            lambda: self.service.sheet_candidates(self.progress.emit),
            self.sheets_ready,
        )

    def analyze_sheet(self, sheet_name: str) -> None:
        self._submit(
            "analysis",
            lambda: self.service.analyze_sheet(
                sheet_name,
                self.progress.emit,
            ),
            self.plan_ready,
        )

    def launch(self, plan: Any, selected_sqt: Iterable[str]) -> None:
        values = tuple(selected_sqt)

        def worker() -> Any:
            prepared = self.service.prepare_selection(
                plan,
                values,
                self.progress.emit,
            )
            self.progress.emit("Đang khởi chạy BAT của PAD…")
            result = self.launcher.launch(prepared)
            record_launched = getattr(self.service, "record_launched", None)
            if callable(record_launched):
                record_launched(prepared)
            return result

        self._submit("launch", worker, self.launched)

    def load_latest_launched(self) -> dict[str, Any] | None:
        loader = getattr(self.service, "load_latest_launched", None)
        return loader() if callable(loader) else None

    def _submit(self, phase: str, worker: Any, success_signal: Signal) -> None:
        with self._lock:
            if self._phase is not None:
                raise RuntimeError("Một tác vụ RPA khác đang được xử lý.")
            self._phase = phase
        self.started.emit(phase)
        try:
            future = self._executor.submit(worker)
        except Exception:
            self._release(phase)
            raise
        future.add_done_callback(
            lambda completed, current=phase, signal=success_signal: (
                self._complete(current, completed, signal)
            )
        )

    def _complete(
        self,
        phase: str,
        future: Future[Any],
        success_signal: Signal,
    ) -> None:
        try:
            result = future.result()
        except Exception as exc:
            self._release(phase)
            self.failed.emit(exc)
            self.finished.emit(phase)
            return
        # Nhả trạng thái trước khi mở dialog ở UI, để người dùng có thể
        # chuyển ngay sang tác vụ phân tích hoặc khởi chạy kế tiếp.
        self._release(phase)
        self.finished.emit(phase)
        success_signal.emit(result)

    def _release(self, phase: str) -> None:
        with self._lock:
            if self._phase == phase:
                self._phase = None

    def shutdown(self, *, wait: bool = True) -> None:
        self._executor.shutdown(wait=wait, cancel_futures=True)


__all__ = ["RpaExpenseController"]
