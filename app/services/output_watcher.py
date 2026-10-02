"""Theo dõi Output và chuyển sự kiện watchdog sang Qt một cách an toàn."""

from __future__ import annotations

import fnmatch
import logging
import os
import threading
import time
from collections.abc import Callable, Iterable, Mapping
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any

from PySide6.QtCore import QObject, Qt, Signal, Slot
from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer

from app.services.file_stability import (
    FileStabilityChecker,
    FileStabilityError,
    FileTooLargeError,
)

LOGGER = logging.getLogger(__name__)

DEFAULT_FILE_PATTERN = "ket_qua_boc_tach*.json"
DEFAULT_MAX_SIZE_BYTES = 50 * 1024 * 1024
TEMPORARY_SUFFIXES = (
    ".crdownload",
    ".part",
    ".partial",
    ".tmp",
    ".temp",
    ".download",
)


class _QtWatchdogBridge(QObject):
    """Signal bridge; watchdog chỉ phát signal, không gọi UI trực tiếp."""

    path_observed = Signal(str)


class _OutputEventHandler(FileSystemEventHandler):
    def __init__(self, bridge: _QtWatchdogBridge) -> None:
        super().__init__()
        self._bridge = bridge

    def on_created(self, event: FileSystemEvent) -> None:
        self._forward(event)

    def on_modified(self, event: FileSystemEvent) -> None:
        self._forward(event)

    def on_closed(self, event: FileSystemEvent) -> None:
        self._forward(event)

    def on_moved(self, event: FileSystemEvent) -> None:
        if event.is_directory:
            return
        destination = getattr(event, "dest_path", "")
        if destination:
            self._bridge.path_observed.emit(os.fsdecode(destination))

    def _forward(self, event: FileSystemEvent) -> None:
        if not event.is_directory:
            self._bridge.path_observed.emit(os.fsdecode(event.src_path))


class OutputWatcher(QObject):
    """Watchdog service phát file JSON sau khi file đã ghi xong.

    Watchdog chạy trong thread riêng. Sự kiện thô được đưa qua
    ``_QtWatchdogBridge``; kiểm tra ổn định chạy trong worker pool; kết quả lại
    được chuyển về thread sở hữu QObject trước khi phát signal công khai hoặc
    gọi ``on_file_ready``.
    """

    file_detected = Signal(str)
    file_ready = Signal(str)
    file_ready_with_origin = Signal(str, bool)
    file_rejected = Signal(str, str)
    watcher_error = Signal(str)
    started = Signal(str)
    stopped = Signal()
    scan_completed = Signal(int)
    status_changed = Signal(bool, str)
    file_processed = Signal(str, object)
    file_processed_with_origin = Signal(str, object, bool)

    _stability_succeeded = Signal(str, object, int, bool)
    _stability_failed = Signal(str, str, bool, int)
    _scan_finished = Signal(int, int)

    def __init__(
        self,
        output_dir: str | Path | object,
        *,
        file_pattern: str | Iterable[str] | None = None,
        pattern: str | Iterable[str] | None = None,
        patterns: str | Iterable[str] | None = None,
        stable_seconds: float | None = None,
        timeout_seconds: float | None = None,
        stability_timeout_seconds: float | None = None,
        poll_interval: float = 0.25,
        max_size_bytes: int | None = None,
        max_file_size_mb: int | float | None = None,
        stability_checker: FileStabilityChecker | None = None,
        on_file_ready: Callable[[Path], object] | None = None,
        callback: Callable[[Path], object] | None = None,
        observer_factory: Callable[[], Any] = Observer,
        max_workers: int = 1,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)

        settings = (
            output_dir
            if not isinstance(output_dir, (str, bytes, os.PathLike))
            else None
        )
        configured_output = self._read_setting(
            settings, "output_dir", default=output_dir
        )
        self._output_dir = Path(os.fsdecode(configured_output)).expanduser()

        configured_pattern = self._read_setting(
            settings,
            "file_pattern",
            "pattern",
            default=DEFAULT_FILE_PATTERN,
        )
        selected_patterns: object = configured_pattern
        for supplied_patterns in (patterns, pattern, file_pattern):
            if supplied_patterns is not None:
                selected_patterns = supplied_patterns
        self._patterns = self._normalize_patterns(selected_patterns)

        configured_stable_seconds = self._read_setting(
            settings, "stable_seconds", default=3.0
        )
        configured_timeout = self._read_setting(
            settings,
            "stability_timeout_seconds",
            "timeout_seconds",
            default=60.0,
        )
        configured_max_mb = self._read_setting(
            settings, "max_file_size_mb", default=50
        )
        effective_stable_seconds = (
            float(stable_seconds)
            if stable_seconds is not None
            else float(configured_stable_seconds)
        )
        supplied_timeout = (
            stability_timeout_seconds
            if stability_timeout_seconds is not None
            else timeout_seconds
        )
        effective_timeout = (
            float(supplied_timeout)
            if supplied_timeout is not None
            else float(configured_timeout)
        )
        if max_size_bytes is not None:
            effective_max_size = int(max_size_bytes)
        elif max_file_size_mb is not None:
            effective_max_size = int(float(max_file_size_mb) * 1024 * 1024)
        else:
            effective_max_size = int(float(configured_max_mb) * 1024 * 1024)

        if effective_stable_seconds < 0:
            raise ValueError("stable_seconds không được âm.")
        if effective_timeout <= 0:
            raise ValueError("timeout_seconds phải lớn hơn 0.")
        if poll_interval <= 0:
            raise ValueError("poll_interval phải lớn hơn 0.")
        if effective_max_size <= 0:
            raise ValueError("Giới hạn dung lượng file phải lớn hơn 0.")
        if max_workers <= 0:
            raise ValueError("max_workers phải lớn hơn 0.")

        self._stable_seconds = effective_stable_seconds
        self._timeout_seconds = effective_timeout
        self._poll_interval = float(poll_interval)
        self._max_size_bytes = effective_max_size
        self._checker = stability_checker or FileStabilityChecker(
            stable_seconds=self._stable_seconds,
            timeout_seconds=self._timeout_seconds,
            poll_interval=self._poll_interval,
            max_size_bytes=effective_max_size,
        )
        self._on_file_ready_callback = on_file_ready or callback
        self._observer_factory = observer_factory
        self._max_workers = max_workers

        self._state_lock = threading.RLock()
        self._observer: Any | None = None
        self._executor: ThreadPoolExecutor | None = None
        self._cancel_event = threading.Event()
        self._pending: set[str] = set()
        self._startup_signatures: dict[str, tuple[int, int]] = {}
        self._handled_signatures: dict[str, tuple[int, int]] = {}
        self._running = False
        self._generation = 0

        self._bridge = _QtWatchdogBridge(self)
        self._event_handler = _OutputEventHandler(self._bridge)
        self._bridge.path_observed.connect(
            self._on_watchdog_path, Qt.ConnectionType.QueuedConnection
        )
        self._stability_succeeded.connect(
            self._deliver_ready, Qt.ConnectionType.QueuedConnection
        )
        self._stability_failed.connect(
            self._deliver_failure, Qt.ConnectionType.QueuedConnection
        )
        self._scan_finished.connect(
            self._deliver_scan_completed, Qt.ConnectionType.QueuedConnection
        )

    @property
    def output_dir(self) -> Path:
        return self._output_dir

    @property
    def file_patterns(self) -> tuple[str, ...]:
        return self._patterns

    @property
    def max_size_bytes(self) -> int:
        return self._max_size_bytes

    @property
    def stable_seconds(self) -> float:
        return self._stable_seconds

    @property
    def timeout_seconds(self) -> float:
        return self._timeout_seconds

    @property
    def is_running(self) -> bool:
        with self._state_lock:
            return self._running

    @property
    def event_handler(self) -> FileSystemEventHandler:
        """Expose handler để unit test có thể phát sự kiện giả."""

        return self._event_handler

    def start(self) -> bool:
        """Khởi động observer và quét file mới nhất đã có trong Output."""

        with self._state_lock:
            if self._running:
                return False

        try:
            self._output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            message = (
                f"Không thể tạo hoặc truy cập thư mục Output: "
                f"{self._output_dir}"
            )
            LOGGER.exception("%s (%s)", message, exc)
            self.watcher_error.emit(message)
            self.status_changed.emit(False, message)
            return False

        cancel_event = threading.Event()
        startup_signatures: dict[str, tuple[int, int]] = {}
        try:
            for path in self._output_dir.iterdir():
                if path.is_file() and self.accepts_path(path):
                    stat_result = path.stat()
                    startup_signatures[self._path_key(path)] = (
                        stat_result.st_size, stat_result.st_mtime_ns
                    )
        except OSError:
            LOGGER.warning("Không thể nhận diện đầy đủ file có sẵn khi khởi động.")
        executor = ThreadPoolExecutor(
            max_workers=self._max_workers,
            thread_name_prefix="output-stability",
        )
        observer = self._observer_factory()

        with self._state_lock:
            self._generation += 1
            generation = self._generation
            self._cancel_event = cancel_event
            self._executor = executor
            self._observer = observer
            self._running = True
            self._pending.clear()
            self._startup_signatures = startup_signatures

        try:
            observer.schedule(
                self._event_handler, str(self._output_dir), recursive=False
            )
            observer.start()
        except Exception as exc:
            with self._state_lock:
                self._running = False
                self._generation += 1
                self._observer = None
                self._executor = None
            cancel_event.set()
            try:
                observer.stop()
                observer.join(timeout=1.0)
            except Exception:
                LOGGER.debug(
                    "Observer không cần hoặc không thể cleanup sau lỗi start.",
                    exc_info=True,
                )
            executor.shutdown(wait=True, cancel_futures=True)
            message = f"Không thể theo dõi thư mục Output: {self._output_dir}"
            LOGGER.exception("%s (%s)", message, exc)
            self.watcher_error.emit(message)
            self.status_changed.emit(False, message)
            return False

        LOGGER.info(
            "Watcher bắt đầu: output=%s, pattern=%s, max_size=%d",
            self._output_dir,
            ", ".join(self._patterns),
            self._max_size_bytes,
        )
        self.started.emit(str(self._output_dir))
        self.status_changed.emit(True, f"Đang theo dõi: {self._output_dir}")
        self.scan_existing(generation=generation)
        return True

    def stop(self, join_timeout: float = 5.0) -> bool:
        """Dừng observer/worker sạch; an toàn khi gọi nhiều lần."""

        if join_timeout <= 0:
            raise ValueError("join_timeout phải lớn hơn 0.")

        with self._state_lock:
            if not self._running and self._observer is None:
                return False
            self._running = False
            self._generation += 1
            observer = self._observer
            executor = self._executor
            cancel_event = self._cancel_event
            self._observer = None
            self._executor = None

        cancel_event.set()
        if observer is not None:
            try:
                observer.stop()
                observer.join(timeout=join_timeout)
                if observer.is_alive():
                    LOGGER.error(
                        "Watchdog chưa dừng sau %.1f giây.", join_timeout
                    )
            except Exception as exc:
                LOGGER.exception("Lỗi khi dừng watchdog: %s", exc)

        if executor is not None:
            executor.shutdown(wait=True, cancel_futures=True)

        with self._state_lock:
            self._pending.clear()
            self._startup_signatures.clear()

        LOGGER.info("Watcher đã dừng: %s", self._output_dir)
        self.stopped.emit()
        self.status_changed.emit(False, "Đã dừng theo dõi thư mục Output.")
        return True

    # Tên gọi phổ biến trong lifecycle của controller.
    close = stop
    shutdown = stop

    def restart(
        self,
        output_dir: str | Path | None = None,
        *,
        file_pattern: str | Iterable[str] | None = None,
    ) -> bool:
        """Dừng watcher cũ và theo dõi thư mục/cấu hình mới."""

        self.stop()
        if output_dir is not None:
            self._output_dir = Path(output_dir).expanduser()
        if file_pattern is not None:
            self._patterns = self._normalize_patterns(file_pattern)
        with self._state_lock:
            self._handled_signatures.clear()
        return self.start()

    def update_settings(
        self,
        settings: object | Mapping[str, Any],
        *,
        restart: bool = True,
    ) -> bool:
        """Áp dụng toàn bộ settings watcher và dựng lại stability checker.

        Nếu watcher đang chạy, observer/worker cũ luôn được dừng trước khi đổi
        cấu hình. ``restart=False`` để watcher ở trạng thái dừng để caller có
        thể chủ động gọi :meth:`start` sau đó.
        """

        new_output = Path(
            os.fsdecode(
                self._read_setting(
                    settings, "output_dir", default=self._output_dir
                )
            )
        ).expanduser()
        new_patterns = self._patterns
        new_stable_seconds = self._stable_seconds
        new_timeout_seconds = self._timeout_seconds
        new_max_size_bytes = self._max_size_bytes

        # Khởi tạo trước để validation thất bại không làm mất watcher hiện tại.
        new_checker = FileStabilityChecker(
            stable_seconds=new_stable_seconds,
            timeout_seconds=new_timeout_seconds,
            poll_interval=self._poll_interval,
            max_size_bytes=new_max_size_bytes,
        )

        self.stop()
        with self._state_lock:
            self._output_dir = new_output
            self._patterns = new_patterns
            self._stable_seconds = new_stable_seconds
            self._timeout_seconds = new_timeout_seconds
            self._max_size_bytes = new_max_size_bytes
            self._checker = new_checker
            self._handled_signatures.clear()

        LOGGER.info(
            "Đã cập nhật watcher: output=%s, pattern=%s, stable=%.3fs, "
            "timeout=%.3fs, max_size=%d",
            self._output_dir,
            ", ".join(self._patterns),
            self._stable_seconds,
            self._timeout_seconds,
            self._max_size_bytes,
        )
        return self.start() if restart else True

    def scan_existing(self, *, generation: int | None = None) -> int:
        """Quét file phù hợp đang có; trả về số candidate đã xếp hàng."""

        with self._state_lock:
            current_generation = self._generation
            running = self._running
        if generation is None:
            generation = current_generation
        if not running or generation != current_generation:
            return 0

        try:
            paths = sorted(
                (
                    path
                    for path in self._output_dir.iterdir()
                    if path.is_file() and self.accepts_path(path)
                ),
                key=lambda path: (path.stat().st_mtime_ns, path.name.casefold()),
                reverse=True,
            )
        except OSError as exc:
            message = f"Không thể quét thư mục Output: {self._output_dir}"
            LOGGER.exception("%s (%s)", message, exc)
            self.watcher_error.emit(message)
            return 0

        queued = 0
        # Chỉ file mới nhất được tiếp nhận; BatchService sẽ dọn các bản cũ.
        for path in paths[:1]:
            if self.enqueue(path, generation=generation):
                queued += 1
        self._scan_finished.emit(queued, generation)
        return queued

    def enqueue(
        self, path: str | Path, *, generation: int | None = None
    ) -> bool:
        """Xếp một candidate vào worker; hữu ích cho sự kiện giả trong test."""

        candidate = Path(path)
        if not self.accepts_path(candidate):
            return False

        try:
            if not candidate.is_file():
                return False
            initial_stat = candidate.stat()
        except OSError:
            return False

        key = self._path_key(candidate)
        signature = (initial_stat.st_size, initial_stat.st_mtime_ns)
        with self._state_lock:
            if not self._running:
                return False
            current_generation = self._generation
            if generation is None:
                generation = current_generation
            if generation != current_generation:
                return False
            if key in self._pending:
                return False
            if self._handled_signatures.get(key) == signature:
                return False
            is_startup = self._startup_signatures.get(key) == signature
            if initial_stat.st_size > self._max_size_bytes:
                self._handled_signatures[key] = signature
                message = (
                    f"File vượt giới hạn "
                    f"{self._format_size(self._max_size_bytes)}."
                )
                LOGGER.warning("%s File=%s", message, candidate)
                self.file_rejected.emit(str(candidate), message)
                return False
            executor = self._executor
            if executor is None:
                return False
            self._pending.add(key)
            cancel_event = self._cancel_event

        LOGGER.info("Phát hiện file JSON: %s", candidate)
        self.file_detected.emit(str(candidate))
        try:
            executor.submit(
                self._wait_for_candidate,
                candidate,
                key,
                generation,
                cancel_event,
                is_startup,
            )
        except RuntimeError as exc:
            with self._state_lock:
                self._pending.discard(key)
            LOGGER.warning("Không thể xếp file vào worker: %s", exc)
            return False
        return True

    # Alias phục vụ controller/test không cần biết tên triển khai.
    process_path = enqueue
    handle_path = enqueue

    def accepts_path(self, path: str | Path) -> bool:
        """Kiểm tra tên/đuôi file và loại các hậu tố tải tạm."""

        candidate = Path(path)
        name = candidate.name
        lowered = name.casefold()
        if not name or lowered.startswith("~$"):
            return False
        if lowered.endswith(TEMPORARY_SUFFIXES):
            return False
        if candidate.suffix.casefold() != ".json":
            return False
        return any(
            fnmatch.fnmatchcase(lowered, pattern.casefold())
            for pattern in self._patterns
        )

    # Tên dễ đọc cho test lọc file.
    is_candidate = accepts_path

    def should_process(self, path: str | Path) -> bool:
        """Kiểm tra đầy đủ tên, file tồn tại và giới hạn dung lượng."""

        candidate = Path(path)
        if not self.accepts_path(candidate):
            return False
        try:
            return (
                candidate.is_file()
                and candidate.stat().st_size <= self._max_size_bytes
            )
        except OSError:
            return False

    def wait_for_idle(self, timeout: float = 10.0) -> bool:
        """Chờ worker hết việc (chủ yếu dành cho test/đóng ứng dụng)."""

        if timeout < 0:
            raise ValueError("timeout không được âm.")
        deadline = time.monotonic() + timeout
        while time.monotonic() <= deadline:
            with self._state_lock:
                if not self._pending:
                    return True
            time.sleep(0.01)
        return False

    @Slot(str)
    def _on_watchdog_path(self, raw_path: str) -> None:
        self.enqueue(Path(raw_path))

    def _wait_for_candidate(
        self,
        path: Path,
        key: str,
        generation: int,
        cancel_event: threading.Event,
        is_startup: bool,
    ) -> None:
        LOGGER.info("Chờ file ổn định: %s", path)
        try:
            result = self._checker.wait(path, cancel_event=cancel_event)
            size, mtime_ns = self._result_signature(result, path)
        except FileTooLargeError as exc:
            signature = self._safe_signature(path)
            with self._state_lock:
                if signature is not None:
                    self._handled_signatures[key] = signature
                self._pending.discard(key)
            message = (
                f"File vượt giới hạn {self._format_size(self._max_size_bytes)}."
            )
            LOGGER.warning("%s File=%s Chi tiết=%s", message, path, exc)
            self._stability_failed.emit(str(path), message, True, generation)
            return
        except FileStabilityError as exc:
            with self._state_lock:
                self._pending.discard(key)
            if cancel_event.is_set():
                LOGGER.debug("Hủy chờ file khi watcher dừng: %s", path)
                return
            message = "File chưa ghi xong hoặc không thể đọc ổn định."
            LOGGER.warning("%s File=%s Chi tiết=%s", message, path, exc)
            self._stability_failed.emit(str(path), message, False, generation)
            return
        except Exception as exc:
            with self._state_lock:
                self._pending.discard(key)
            if cancel_event.is_set():
                return
            message = "Không thể kiểm tra file vừa tải."
            LOGGER.exception("%s File=%s Chi tiết=%s", message, path, exc)
            self._stability_failed.emit(str(path), message, False, generation)
            return

        with self._state_lock:
            self._handled_signatures[key] = (size, mtime_ns)
            self._pending.discard(key)
            is_startup = is_startup and self._startup_signatures.get(key) == (
                size, mtime_ns
            )
        self._stability_succeeded.emit(str(path), result, generation, is_startup)

    @Slot(str, object, int, bool)
    def _deliver_ready(
        self, raw_path: str, result: object, generation: int, is_startup: bool = False
    ) -> None:
        if not self._is_current_generation(generation):
            return
        path = Path(raw_path)
        latest = self._latest_candidate()
        if latest is None or self._path_key(latest) != self._path_key(path):
            LOGGER.info(
                "Bỏ qua candidate cũ vì Output đã có file mới hơn: %s",
                path,
            )
            return
        LOGGER.info("File đã ổn định và sẵn sàng tiếp nhận: %s", path)
        self.file_ready_with_origin.emit(raw_path, is_startup)
        self.file_ready.emit(raw_path)
        if self._on_file_ready_callback is None:
            return
        try:
            callback_result = self._on_file_ready_callback(path)
        except Exception as exc:
            message = (
                "Không thể tiếp nhận file JSON. Hãy xem Nhật ký để biết chi tiết."
            )
            LOGGER.exception(
                "Callback tiếp nhận file thất bại, file=%s: %s", path, exc
            )
            self.file_rejected.emit(raw_path, message)
            return
        self.file_processed_with_origin.emit(raw_path, callback_result, is_startup)
        self.file_processed.emit(raw_path, callback_result)

    @Slot(str, str, bool, int)
    def _deliver_failure(
        self,
        raw_path: str,
        message: str,
        _rejected: bool,
        generation: int,
    ) -> None:
        if not self._is_current_generation(generation):
            return
        # Đây là lỗi của riêng candidate; observer vẫn đang hoạt động. Chỉ lỗi
        # lifecycle của watcher mới phát watcher_error/status false.
        self.file_rejected.emit(raw_path, message)

    @Slot(int, int)
    def _deliver_scan_completed(self, count: int, generation: int) -> None:
        if self._is_current_generation(generation):
            LOGGER.info("Quét Output hoàn tất: %d file được xếp hàng.", count)
            self.scan_completed.emit(count)

    def _is_current_generation(self, generation: int) -> bool:
        with self._state_lock:
            return self._running and generation == self._generation

    @staticmethod
    def _normalize_patterns(
        patterns: str | Iterable[str] | object,
    ) -> tuple[str, ...]:
        if isinstance(patterns, str):
            raw_patterns = patterns.replace(",", ";").split(";")
        elif isinstance(patterns, Iterable):
            raw_patterns = [str(pattern) for pattern in patterns]
        else:
            raw_patterns = [DEFAULT_FILE_PATTERN]

        normalized = tuple(
            pattern.strip()
            for pattern in raw_patterns
            if pattern and pattern.strip()
        )
        return normalized or (DEFAULT_FILE_PATTERN,)

    @staticmethod
    def _read_setting(
        settings: object | None,
        *names: str,
        default: object,
    ) -> object:
        if settings is None:
            return default
        if isinstance(settings, Mapping):
            for name in names:
                value = settings.get(name)
                if value is not None:
                    return value
            return default
        for name in names:
            value = getattr(settings, name, None)
            if value is not None:
                return value
        return default

    @staticmethod
    def _path_key(path: Path) -> str:
        return os.path.normcase(os.path.abspath(os.fsdecode(path)))

    @staticmethod
    def _result_signature(result: object, path: Path) -> tuple[int, int]:
        size = getattr(result, "size", None)
        mtime_ns = getattr(result, "mtime_ns", None)
        if isinstance(size, int) and isinstance(mtime_ns, int):
            return size, mtime_ns
        stat_result = path.stat()
        return stat_result.st_size, stat_result.st_mtime_ns

    @staticmethod
    def _safe_signature(path: Path) -> tuple[int, int] | None:
        try:
            stat_result = path.stat()
        except OSError:
            return None
        return stat_result.st_size, stat_result.st_mtime_ns

    def _latest_candidate(self) -> Path | None:
        try:
            candidates = [
                path
                for path in self._output_dir.iterdir()
                if path.is_file() and self.accepts_path(path)
            ]
            return max(
                candidates,
                key=lambda path: (
                    path.stat().st_mtime_ns,
                    path.name.casefold(),
                ),
                default=None,
            )
        except OSError:
            return None

    @staticmethod
    def _format_size(size: int) -> str:
        return f"{size / (1024 * 1024):g} MB"


    def mark_handled(self, path: str | Path) -> None:
        """Bỏ qua đúng phiên bản file vừa được chính ứng dụng ghi."""

        candidate = Path(path)
        signature = self._safe_signature(candidate)
        if signature is None:
            return
        with self._state_lock:
            self._handled_signatures[self._path_key(candidate)] = signature


OutputWatcherService = OutputWatcher

__all__ = [
    "DEFAULT_FILE_PATTERN",
    "DEFAULT_MAX_SIZE_BYTES",
    "OutputWatcher",
    "OutputWatcherService",
    "TEMPORARY_SUFFIXES",
]
