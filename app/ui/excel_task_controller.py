"""Điều phối tác vụ Excel nền qua một worker duy nhất.

Controller không phụ thuộc model/service cụ thể. Service chỉ cần cung cấp
``analyze(progress_callback=...)`` và ``apply(plan, resolutions,
progress_callback=...)``.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from concurrent.futures import Future, ThreadPoolExecutor
from threading import RLock
from typing import Any

from PySide6.QtCore import QObject, Signal

from app.services.excel.review import (
    CorrectionRequiredError,
    ReviewOutcome,
    validate_conflict_resolutions,
)


def _value(source: Any, *names: str, default: Any = None) -> Any:
    if source is None:
        return default
    if isinstance(source, Mapping):
        for name in names:
            if name in source:
                return source[name]
        return default
    for name in names:
        if hasattr(source, name):
            value = getattr(source, name)
            if callable(value) and name.startswith(("is_", "has_", "needs_", "requires_")):
                value = value()
            return getattr(value, "value", value)
    return default


def _items(value: Any) -> tuple[Any, ...]:
    if value is None:
        return ()
    if isinstance(value, Mapping):
        return tuple(value.values())
    if isinstance(value, Sequence) and not isinstance(
        value, (str, bytes, bytearray)
    ):
        return tuple(value)
    try:
        return tuple(value)
    except TypeError:
        return (value,)


class ExcelTaskController(QObject):
    """Chạy đồng bộ/nhập khoản chi tuần tự và chuyển kết quả về Qt UI.

    Signal công khai:

    - ``started(str operation)``
    - ``progress(str operation, str message)``
    - ``analysis_ready(object plan)``
    - ``correction_required(object ReviewOutcome)``
    - ``completed(object result)``
    - ``failed(object exception)``
    - ``finished(str operation)``
    """

    started = Signal(str)
    progress = Signal(str, str)
    analysis_ready = Signal(object)
    correction_required = Signal(object)
    completed = Signal(object)
    failed = Signal(object)
    finished = Signal(str)

    SYNC_OPERATION = "sync"
    POSTING_OPERATION = "posting"
    PAYMENT_SYNC_OPERATION = "payment_sync"

    def __init__(
        self,
        daily_sync_service: Any | None = None,
        expense_posting_service: Any | None = None,
        payment_sync_service: Any | None = None,
        draft_service: Any | None = None,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self.daily_sync_service = daily_sync_service
        self.expense_posting_service = expense_posting_service
        self.payment_sync_service = payment_sync_service
        self.draft_service = draft_service
        self._executor = ThreadPoolExecutor(
            max_workers=1,
            thread_name_prefix="excel-worker",
        )
        self._state_lock = RLock()
        self._active_operation: str | None = None
        self._phase = "idle"
        self._future: Future[Any] | None = None
        self._waiting_plan: Any | None = None
        self._base_plan: Any | None = None
        self._review_resolutions: dict[str, Any] = {}
        self._review_conflicts: list[Any] = []
        self._draft_context_key: str | None = None
        self._draft_saved_this_run = False
        self._saved_resolution_info: dict[str, Any] = {}
        self._closed = False

    @classmethod
    def normalize_operation(cls, operation: Any) -> str:
        value = str(getattr(operation, "value", operation) or "").casefold()
        if value in {"sync", "daily_sync", "sync_daily", "daily"}:
            return cls.SYNC_OPERATION
        if value in {
            "posting",
            "post",
            "expense_posting",
            "post_expenses",
            "expenses",
        }:
            return cls.POSTING_OPERATION
        if value in {
            "payment_sync",
            "sync_payment",
            "bk_to_payment",
            "payment",
        }:
            return cls.PAYMENT_SYNC_OPERATION
        raise ValueError(f"Nghiệp vụ Excel không hợp lệ: {operation!r}")

    @property
    def active_operation(self) -> str | None:
        with self._state_lock:
            return self._active_operation

    @property
    def is_busy(self) -> bool:
        return self.active_operation is not None

    @property
    def phase(self) -> str:
        with self._state_lock:
            return self._phase

    def update_services(
        self,
        *,
        daily_sync_service: Any | None = None,
        expense_posting_service: Any | None = None,
        payment_sync_service: Any | None = None,
    ) -> None:
        """Thay service sau khi settings/runtime được cập nhật."""

        if self.is_busy:
            raise RuntimeError("Không thể thay dịch vụ khi tác vụ Excel đang chạy.")
        if daily_sync_service is not None:
            self.daily_sync_service = daily_sync_service
        if expense_posting_service is not None:
            self.expense_posting_service = expense_posting_service
        if payment_sync_service is not None:
            self.payment_sync_service = payment_sync_service

    def submit(
        self,
        operation: Any,
        task: Callable[..., Any],
        *args: Any,
        with_progress: bool = False,
        **kwargs: Any,
    ) -> Future[Any]:
        """Chạy một callable độc quyền trên Excel worker.

        Khi ``with_progress=True``, controller tự truyền keyword
        ``progress_callback`` nếu caller chưa truyền.
        """

        normalized = self.normalize_operation(operation)
        self._reserve(normalized, "running")
        self.started.emit(normalized)
        call_kwargs = dict(kwargs)
        if with_progress:
            call_kwargs.setdefault(
                "progress_callback",
                self._progress_callback(normalized),
            )
        try:
            future = self._executor.submit(task, *args, **call_kwargs)
        except BaseException:
            self._release(normalized)
            self.finished.emit(normalized)
            raise
        self._future = future
        future.add_done_callback(
            lambda done, op=normalized: self._complete_one_shot(op, done)
        )
        return future

    def start_sync(self, **analyze_kwargs: Any) -> Future[Any]:
        return self._start_analysis(
            self.SYNC_OPERATION,
            self.daily_sync_service,
            analyze_kwargs,
        )

    analyze_sync = start_sync
    submit_sync = start_sync

    def start_posting(self, **analyze_kwargs: Any) -> Future[Any]:
        return self._start_analysis(
            self.POSTING_OPERATION,
            self.expense_posting_service,
            analyze_kwargs,
        )

    analyze_posting = start_posting
    submit_posting = start_posting

    def start_payment_sync(self, **analyze_kwargs: Any) -> Future[Any]:
        return self._start_analysis(
            self.PAYMENT_SYNC_OPERATION,
            self.payment_sync_service,
            analyze_kwargs,
        )

    analyze_payment_sync = start_payment_sync
    submit_payment_sync = start_payment_sync

    def start(self, operation: Any, **analyze_kwargs: Any) -> Future[Any]:
        normalized = self.normalize_operation(operation)
        if normalized == self.SYNC_OPERATION:
            return self.start_sync(**analyze_kwargs)
        if normalized == self.POSTING_OPERATION:
            return self.start_posting(**analyze_kwargs)
        return self.start_payment_sync(**analyze_kwargs)

    def _start_analysis(
        self,
        operation: str,
        service: Any,
        analyze_kwargs: Mapping[str, Any],
    ) -> Future[Any]:
        if service is None or not callable(getattr(service, "analyze", None)):
            raise RuntimeError(
                "Dịch vụ đồng bộ Excel chưa được khởi tạo."
                if operation == self.SYNC_OPERATION
                else "Dịch vụ nhập khoản chi chưa được khởi tạo."
            )
        self._reserve(operation, "analyzing")
        self.started.emit(operation)
        kwargs = dict(analyze_kwargs)
        kwargs.setdefault("progress_callback", self._progress_callback(operation))
        try:
            future = self._executor.submit(service.analyze, **kwargs)
        except BaseException:
            self._release(operation)
            self.finished.emit(operation)
            raise
        self._future = future
        future.add_done_callback(
            lambda done, op=operation, owner=service: self._analysis_done(
                op, owner, done
            )
        )
        return future

    def apply_plan(
        self,
        plan: Any,
        resolutions: Any = None,
        *,
        operation: Any | None = None,
    ) -> Future[Any]:
        """Tiếp tục một plan đang chờ dialog hoặc áp dụng plan do caller cung cấp."""

        normalized = (
            self.normalize_operation(operation)
            if operation is not None
            else self._operation_for_plan(plan)
        )
        service = self._service_for(normalized)
        if service is None or not callable(getattr(service, "apply", None)):
            raise RuntimeError("Dịch vụ Excel không hỗ trợ áp dụng kế hoạch.")
        self.save_draft(plan, resolutions, operation=normalized)

        with self._state_lock:
            active = self._active_operation
            phase = self._phase
            if active is None:
                self._active_operation = normalized
                self._phase = "applying"
                emit_started = True
            elif active != normalized:
                raise RuntimeError(
                    f"Tác vụ {active!r} đang hoạt động; không thể chạy {normalized!r}."
                )
            elif phase != "waiting_user":
                raise RuntimeError("Kế hoạch Excel chưa ở trạng thái chờ xử lý.")
            else:
                self._phase = "applying"
                emit_started = False
            self._waiting_plan = None

        if emit_started:
            self.started.emit(normalized)
        return self._submit_apply(
            normalized,
            service,
            plan,
            {} if resolutions is None else resolutions,
        )

    continue_with_resolutions = apply_plan
    apply = apply_plan

    def saved_resolutions(
        self,
        plan: Any,
        *,
        operation: Any | None = None,
    ) -> dict[str, Any]:
        """Return compatible choices from the unfinished session, if any."""

        if self.draft_service is None:
            self._saved_resolution_info = {}
            return {}
        normalized = (
            self.normalize_operation(operation)
            if operation is not None
            else self._operation_for_plan(plan)
        )
        restore_with_info = getattr(self.draft_service, "restore_with_info", None)
        if callable(restore_with_info):
            restored = restore_with_info(plan, normalized)
            self._draft_context_key = str(restored.source_file_key)
            self._saved_resolution_info = {
                "found": bool(restored.found),
                "status": restored.status,
                "updated_at": restored.updated_at,
                "current_conflict_count": restored.current_conflict_count,
                "saved_choice_count": restored.saved_choice_count,
                "restored_count": restored.restored_count,
                "target_compatible": restored.target_compatible,
                "last_error": restored.last_error,
            }
            return dict(restored.resolutions)
        context_key, resolutions = self.draft_service.restore(plan, normalized)
        self._draft_context_key = context_key
        self._saved_resolution_info = {}
        return dict(resolutions)

    @property
    def saved_resolution_info(self) -> dict[str, Any]:
        return dict(self._saved_resolution_info)

    def save_draft(
        self,
        plan: Any,
        resolutions: Mapping[str, Any] | None,
        *,
        operation: Any | None = None,
    ) -> None:
        """Persist choices before refine/apply so failures do not discard them."""

        if self.draft_service is None:
            return
        normalized = (
            self.normalize_operation(operation)
            if operation is not None
            else self._operation_for_plan(plan)
        )
        active_conflicts = tuple(self._review_conflicts)
        try:
            self._draft_context_key = self.draft_service.save(
                plan,
                normalized,
                resolutions,
                conflicts=(active_conflicts or None),
            )
        except TypeError:
            # Tương thích các draft adapter cũ/custom trong quá trình nâng cấp.
            self._draft_context_key = self.draft_service.save(
                plan, normalized, resolutions
            )
        self._draft_saved_this_run = True

    def refine_plan(
        self,
        plan: Any,
        resolutions: Any,
        *,
        operation: Any | None = None,
    ) -> Future[Any]:
        """Compatibility alias for the full base-plan review preparation."""

        return self.prepare_plan(
            plan,
            resolutions,
            operation=operation,
        )

    def prepare_plan(
        self,
        plan: Any,
        resolutions: Any,
        *,
        operation: Any | None = None,
    ) -> Future[Any]:
        """Replay every choice from the base plan before any workbook write."""

        normalized = (
            self.normalize_operation(operation)
            if operation is not None
            else self._operation_for_plan(plan)
        )
        service = self._service_for(normalized)
        self.save_draft(plan, resolutions, operation=normalized)
        with self._state_lock:
            if (
                self._active_operation != normalized
                or self._phase != "waiting_user"
            ):
                raise RuntimeError("Kế hoạch Excel chưa ở trạng thái chờ xử lý.")
            self._phase = "preparing_review"
            self._waiting_plan = None
            base_plan = self._base_plan or plan
            self._review_resolutions = dict(resolutions or {})
        callback = self._progress_callback(normalized)
        try:
            future = self._executor.submit(
                self._prepare_review,
                service,
                base_plan,
                dict(resolutions or {}),
                callback,
            )
        except BaseException as exc:
            self._finish_failure(normalized, exc)
            raise
        self._future = future
        future.add_done_callback(
            lambda done, op=normalized, owner=service: self._refine_done(
                op, owner, done
            )
        )
        return future

    def _prepare_review(
        self,
        service: Any,
        base_plan: Any,
        resolutions: Mapping[str, Any],
        progress_callback: Callable[..., None],
    ) -> ReviewOutcome:
        current = base_plan
        ledger = dict(resolutions)
        seen_states: set[tuple[str, ...]] = set()
        refine = getattr(service, "refine", None)

        for _stage in range(32):
            conflicts = list(
                _items(_value(current, "conflicts", "unresolved_conflicts", default=()))
            )
            restore_dynamic = getattr(
                self.draft_service, "restore_for_conflicts", None
            )
            if conflicts and callable(restore_dynamic):
                restored = restore_dynamic(
                    base_plan,
                    self._operation_for_plan(base_plan),
                    conflicts,
                )
                for conflict_id, value in dict(restored or {}).items():
                    ledger.setdefault(str(conflict_id), value)
            # Mỗi lần refine là một bước quyết định mới. Chỉ đưa các conflict
            # hiện tại về dialog để lựa chọn nguồn đã xử lý không xuất hiện
            # ngang cấp với conflict ô đích vừa được phát hiện.
            issues = validate_conflict_resolutions(conflicts, ledger)
            if issues:
                return ReviewOutcome.needs_correction(
                    conflicts=conflicts,
                    issues=issues,
                    resolutions=ledger,
                )
            if not conflicts or not callable(refine):
                return ReviewOutcome(
                    prepared_plan=current,
                    conflicts=conflicts,
                    resolutions=dict(ledger),
                )

            state = tuple(
                str(_value(conflict, "conflict_id", "id", default=""))
                for conflict in conflicts
            )
            if state in seen_states:
                return ReviewOutcome.needs_correction(
                    conflicts=conflicts,
                    resolutions=ledger,
                )
            seen_states.add(state)
            progress_callback("Đang kiểm tra lại toàn bộ lựa chọn xung đột…")
            # Mỗi vòng refine chỉ được tiêu thụ quyết định của các conflict đang
            # hiện diện. Nếu truyền cả ledger của phiên, một quyết định dành cho
            # conflict động ở vòng sau có thể làm conflict đó biến mất trước khi
            # service kịp áp dụng quyết định vào item.
            current_ids = {
                str(_value(conflict, "conflict_id", "id", default=""))
                for conflict in conflicts
            }
            stage_resolutions = {
                conflict_id: value
                for conflict_id, value in ledger.items()
                if str(conflict_id) in current_ids
            }
            current = refine(
                current,
                stage_resolutions,
                progress_callback=progress_callback,
            )

        remaining = list(
            _items(_value(current, "conflicts", "unresolved_conflicts", default=()))
        )
        return ReviewOutcome.needs_correction(
            conflicts=remaining,
            resolutions=ledger,
        )

    def cancel_waiting(self) -> bool:
        """Giải phóng controller khi người dùng đóng dialog xung đột."""

        with self._state_lock:
            if self._active_operation is None or self._phase != "waiting_user":
                return False
            operation = self._active_operation
            plan = self._waiting_plan
        service = self._service_for(operation)
        cancel = getattr(service, "cancel", None)
        if plan is not None and callable(cancel):
            try:
                cancel(plan)
            except Exception:
                # Hủy UI không được làm controller mắc kẹt chỉ vì lưu audit lỗi.
                pass
        if self.draft_service is not None and self._draft_saved_this_run:
            self.draft_service.mark_cancelled(self._draft_context_key)
        self._release(operation)
        self.finished.emit(operation)
        return True

    def _analysis_done(
        self,
        operation: str,
        service: Any,
        future: Future[Any],
    ) -> None:
        try:
            plan = future.result()
        except BaseException as exc:
            self._finish_failure(operation, exc)
            return

        if self._requires_user_input(plan):
            with self._state_lock:
                if self._active_operation == operation:
                    self._phase = "waiting_user"
                    self._waiting_plan = plan
                    self._base_plan = plan
                    self._review_resolutions = {}
                    self._review_conflicts = list(
                        _items(
                            _value(
                                plan,
                                "conflicts",
                                "unresolved_conflicts",
                                default=(),
                            )
                        )
                    )
            self.analysis_ready.emit(plan)
            return

        if self._analysis_is_terminal(plan):
            apply_method = getattr(service, "apply", None)
            if callable(apply_method):
                self._submit_apply(operation, service, plan, {})
            else:
                self._finish_success(operation, plan)
            return

        self._submit_apply(operation, service, plan, {})

    def _refine_done(
        self,
        operation: str,
        service: Any,
        future: Future[Any],
    ) -> None:
        try:
            outcome = future.result()
        except CorrectionRequiredError as exc:
            self._return_to_review(
                operation,
                ReviewOutcome.needs_correction(
                    conflicts=self._review_conflicts,
                    issues=exc.issues,
                    resolutions=self._review_resolutions,
                ),
            )
            return
        except BaseException as exc:
            self._finish_failure(operation, exc)
            return
        if not isinstance(outcome, ReviewOutcome):
            outcome = ReviewOutcome(prepared_plan=outcome)
        if not outcome.ready:
            self._return_to_review(operation, outcome)
            return

        plan = outcome.prepared_plan
        with self._state_lock:
            self._review_conflicts = list(outcome.conflicts)
        refine = getattr(service, "refine", None)
        if callable(refine) and self._requires_user_input(plan):
            with self._state_lock:
                if self._active_operation == operation:
                    self._phase = "waiting_user"
                    self._waiting_plan = plan
            self.analysis_ready.emit(plan)
            return
        self._submit_apply(
            operation,
            service,
            plan,
            outcome.resolutions if not callable(refine) else {},
        )

    def _return_to_review(
        self, operation: str, outcome: ReviewOutcome
    ) -> None:
        with self._state_lock:
            if self._active_operation != operation:
                return
            self._phase = "waiting_user"
            self._waiting_plan = self._base_plan
            self._review_resolutions = dict(outcome.resolutions)
            self._review_conflicts = list(outcome.conflicts)
        self.correction_required.emit(outcome)

    def _submit_apply(
        self,
        operation: str,
        service: Any,
        plan: Any,
        resolutions: Any,
    ) -> Future[Any]:
        callback = self._progress_callback(operation)
        try:
            future = self._executor.submit(
                service.apply,
                plan,
                resolutions,
                progress_callback=callback,
            )
        except BaseException as exc:
            self._finish_failure(operation, exc)
            raise
        self._future = future
        future.add_done_callback(
            lambda done, op=operation, owner=service, current_plan=plan, current_resolutions=resolutions: self._apply_done(
                op,
                owner,
                current_plan,
                current_resolutions,
                done,
            )
        )
        return future

    def _apply_done(
        self,
        operation: str,
        service: Any,
        plan: Any,
        resolutions: Any,
        future: Future[Any],
    ) -> None:
        try:
            result = future.result()
        except CorrectionRequiredError as exc:
            self._return_to_review(
                operation,
                ReviewOutcome.needs_correction(
                    conflicts=self._review_conflicts
                    or list(
                        _items(
                            _value(
                                self._base_plan or plan,
                                "conflicts",
                                default=(),
                            )
                        )
                    ),
                    issues=exc.issues,
                    resolutions=(
                        resolutions
                        if isinstance(resolutions, Mapping)
                        else self._review_resolutions
                    ),
                ),
            )
            return
        except BaseException as exc:
            self._finish_failure(operation, exc)
            return
        self._finish_success(operation, result)

    def _complete_one_shot(
        self,
        operation: str,
        future: Future[Any],
    ) -> None:
        try:
            result = future.result()
        except BaseException as exc:
            self._finish_failure(operation, exc)
            return
        self._finish_success(operation, result)

    def _finish_success(self, operation: str, result: Any) -> None:
        if self.draft_service is not None and self._draft_saved_this_run:
            self.draft_service.mark_completed(self._draft_context_key)
        self._release(operation)
        self.completed.emit(result)
        self.finished.emit(operation)

    def _finish_failure(self, operation: str, error: BaseException) -> None:
        if self.draft_service is not None and self._draft_saved_this_run:
            self.draft_service.mark_failed(self._draft_context_key, error)
        self._release(operation)
        self.failed.emit(error)
        self.finished.emit(operation)

    def _reserve(self, operation: str, phase: str) -> None:
        with self._state_lock:
            if self._closed:
                raise RuntimeError("ExcelTaskController đã đóng.")
            if self._active_operation is not None:
                raise RuntimeError(
                    f"Tác vụ {self._active_operation!r} đang hoạt động."
                )
            self._active_operation = operation
            self._phase = phase
            self._draft_context_key = None
            self._draft_saved_this_run = False
            self._saved_resolution_info = {}

    def _release(self, operation: str) -> None:
        with self._state_lock:
            if self._active_operation == operation:
                self._active_operation = None
                self._phase = "idle"
                self._future = None
                self._waiting_plan = None
                self._base_plan = None
                self._review_resolutions = {}
                self._review_conflicts = []
                self._draft_context_key = None
                self._draft_saved_this_run = False

    def _progress_callback(self, operation: str) -> Callable[..., None]:
        def report(*parts: Any) -> None:
            if not parts:
                return
            message = " ".join(str(part) for part in parts if part is not None)
            if message:
                self.progress.emit(operation, message)

        return report

    def _service_for(self, operation: str) -> Any:
        if operation == self.SYNC_OPERATION:
            return self.daily_sync_service
        if operation == self.POSTING_OPERATION:
            return self.expense_posting_service
        return self.payment_sync_service

    def _operation_for_plan(self, plan: Any) -> str:
        explicit = _value(plan, "operation", "operation_type")
        if explicit:
            return self.normalize_operation(explicit)
        name = type(plan).__name__.casefold()
        if "post" in name or "expense" in name:
            return self.POSTING_OPERATION
        if "payment" in name:
            return self.PAYMENT_SYNC_OPERATION
        if "sync" in name:
            return self.SYNC_OPERATION
        active = self.active_operation
        if active is not None:
            return active
        raise ValueError("Không xác định được nghiệp vụ từ kế hoạch Excel.")

    @staticmethod
    def _requires_user_input(plan: Any) -> bool:
        explicit = _value(
            plan,
            "requires_user_input",
            "needs_user_input",
            "needs_resolution",
            default=None,
        )
        if explicit is True:
            return bool(explicit)

        conflicts = _items(
            _value(plan, "conflicts", "unresolved_conflicts", default=())
        )
        if conflicts:
            return True
        if explicit is False:
            return False

        selected = _value(
            plan,
            "selected_sheet_name",
            "sheet_name",
            "selected_month",
            default=None,
        )
        candidates = _items(
            _value(
                plan,
                "month_candidates",
                "sheet_candidates",
                "target_sheet_candidates",
                default=(),
            )
        )
        return selected in (None, "") and len(candidates) > 1

    @staticmethod
    def _analysis_is_terminal(plan: Any) -> bool:
        explicit = _value(
            plan,
            "is_terminal",
            "analysis_complete",
            default=None,
        )
        if explicit is not None:
            return bool(explicit)
        status = str(_value(plan, "status", default="")).split(".")[-1].upper()
        if status in {"NO_CHANGES", "CANCELLED", "FAILED"}:
            return True
        has_changes = _value(plan, "has_changes", default=None)
        return has_changes is False

    def shutdown(
        self,
        wait: bool = True,
        *,
        cancel_futures: bool = False,
    ) -> None:
        with self._state_lock:
            if self._closed:
                return
            self._closed = True
        self._executor.shutdown(wait=wait, cancel_futures=cancel_futures)

    close = shutdown


__all__ = ["ExcelTaskController"]
