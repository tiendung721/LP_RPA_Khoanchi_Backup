from __future__ import annotations

from pathlib import Path
import os
from types import SimpleNamespace
from typing import Any

from watchdog.events import FileCreatedEvent

from app.services.output_watcher import OutputWatcher


class FakeObserver:
    def __init__(self) -> None:
        self.handler: Any | None = None
        self.directory = ""
        self.running = False

    def schedule(
        self, handler: Any, directory: str, *, recursive: bool
    ) -> object:
        assert recursive is False
        self.handler = handler
        self.directory = directory
        return object()

    def start(self) -> None:
        self.running = True

    def stop(self) -> None:
        self.running = False

    def join(self, timeout: float | None = None) -> None:
        assert timeout is None or timeout > 0

    def is_alive(self) -> bool:
        return self.running


def make_watcher(
    output: Path,
    observer: FakeObserver,
    *,
    max_size_bytes: int = 1024,
) -> OutputWatcher:
    return OutputWatcher(
        output,
        file_pattern="ket_qua_boc_tach*.json",
        stable_seconds=0,
        timeout_seconds=1,
        poll_interval=0.01,
        max_size_bytes=max_size_bytes,
        observer_factory=lambda: observer,
    )


def test_filter_ignores_temporary_and_wrong_pattern(tmp_path: Path) -> None:
    watcher = OutputWatcher(tmp_path)

    assert watcher.accepts_path("ket_qua_boc_tach.json")
    assert watcher.accepts_path("KET_QUA_BOC_TACH (1).JSON")
    assert watcher.accepts_path("ket_qua_boc_tach_phieu_01.json")
    assert not watcher.accepts_path("ket_qua_boc_tach.json.crdownload")
    assert not watcher.accepts_path("ket_qua_boc_tach.json.part")
    assert not watcher.accepts_path("~$ket_qua_boc_tach.json")
    assert not watcher.accepts_path("ket_qua_khac.json")
    assert not watcher.accepts_path("ket_qua_boc_tach.txt")


def test_startup_scan_waits_for_stable_file_and_emits_on_qt_thread(
    tmp_path: Path,
    qtbot: Any,
) -> None:
    source = tmp_path / "ket_qua_boc_tach.json"
    source.write_text('{"v":1,"d":[]}', encoding="utf-8")
    observer = FakeObserver()
    watcher = make_watcher(tmp_path, observer)

    try:
        with qtbot.waitSignal(watcher.file_ready, timeout=2_000) as signal:
            assert watcher.start() is True
        assert Path(signal.args[0]) == source
        assert observer.directory == str(tmp_path)
        assert observer.running is True
    finally:
        watcher.stop()

    assert observer.running is False
    assert watcher.is_running is False


def test_file_replaced_during_startup_scan_is_treated_as_new(
    tmp_path: Path, qtbot: Any
) -> None:
    source = tmp_path / "ket_qua_boc_tach.json"
    source.write_text('{"v":1,"d":[]}', encoding="utf-8")

    class ReplacingChecker:
        def wait(self, path: Path, *, cancel_event: object) -> SimpleNamespace:
            path.write_text('{"v":1,"d":[["new"]]}', encoding="utf-8")
            stat_result = path.stat()
            return SimpleNamespace(
                size=stat_result.st_size, mtime_ns=stat_result.st_mtime_ns
            )

    watcher = OutputWatcher(
        tmp_path,
        stability_checker=ReplacingChecker(),
        on_file_ready=lambda path: path,
        observer_factory=FakeObserver,
    )
    try:
        with qtbot.waitSignal(watcher.file_processed_with_origin, timeout=2_000) as signal:
            watcher.start()
        assert Path(signal.args[0]) == source
        assert signal.args[2] is False
    finally:
        watcher.stop()


def test_fake_watchdog_event_is_forwarded_and_duplicate_event_is_coalesced(
    tmp_path: Path,
    qtbot: Any,
) -> None:
    observer = FakeObserver()
    watcher = make_watcher(tmp_path, observer)
    ready_paths: list[str] = []
    watcher.file_ready.connect(ready_paths.append)

    try:
        assert watcher.start() is True
        source = tmp_path / "ket_qua_boc_tach_01.json"
        source.write_text('{"v":1,"d":[]}', encoding="utf-8")
        assert observer.handler is not None

        observer.handler.on_created(FileCreatedEvent(str(source)))
        observer.handler.on_created(FileCreatedEvent(str(source)))
        qtbot.waitUntil(lambda: len(ready_paths) == 1, timeout=2_000)
        qtbot.wait(50)
        assert ready_paths == [str(source)]
    finally:
        watcher.stop()


def test_startup_scan_only_queues_newest_matching_file(
    tmp_path: Path,
    qtbot: Any,
) -> None:
    older = tmp_path / "ket_qua_boc_tach.json"
    newer = tmp_path / "ket_qua_boc_tach (1).json"
    older.write_text('{"v":1,"d":[]}', encoding="utf-8")
    newer.write_text('{"v":1,"d":[]}', encoding="utf-8")
    os.utime(older, ns=(1_000_000_000, 1_000_000_000))
    os.utime(newer, ns=(2_000_000_000, 2_000_000_000))
    observer = FakeObserver()
    watcher = make_watcher(tmp_path, observer)

    try:
        with qtbot.waitSignal(watcher.file_ready, timeout=2_000) as signal:
            watcher.start()
        assert Path(signal.args[0]) == newer
    finally:
        watcher.stop()


def test_ready_event_for_older_file_is_ignored_when_newer_exists(
    tmp_path: Path,
) -> None:
    older = tmp_path / "ket_qua_boc_tach.json"
    newer = tmp_path / "ket_qua_boc_tach (1).json"
    older.write_text('{"v":1,"d":[]}', encoding="utf-8")
    newer.write_text('{"v":1,"d":[]}', encoding="utf-8")
    os.utime(older, ns=(1_000_000_000, 1_000_000_000))
    os.utime(newer, ns=(2_000_000_000, 2_000_000_000))
    observer = FakeObserver()
    processed: list[Path] = []
    watcher = OutputWatcher(
        tmp_path,
        stable_seconds=0,
        timeout_seconds=1,
        poll_interval=0.01,
        observer_factory=lambda: observer,
        on_file_ready=lambda path: processed.append(path),
    )

    try:
        watcher.start()
        generation = watcher._generation
        watcher._deliver_ready(str(older), object(), generation)
        watcher._deliver_ready(str(newer), object(), generation)
    finally:
        watcher.stop()

    assert processed == [newer]


def test_mark_handled_prevents_app_write_from_becoming_new_batch(
    tmp_path: Path,
) -> None:
    observer = FakeObserver()
    watcher = make_watcher(tmp_path, observer)

    try:
        watcher.start()
        source = tmp_path / "ket_qua_boc_tach.json"
        source.write_text('{"v":1,"d":[]}', encoding="utf-8")
        watcher.mark_handled(source)
        assert watcher.enqueue(source) is False
    finally:
        watcher.stop()


def test_startup_scan_rejects_file_over_size_limit(
    tmp_path: Path,
    qtbot: Any,
) -> None:
    source = tmp_path / "ket_qua_boc_tach.json"
    source.write_bytes(b"1234")
    observer = FakeObserver()
    watcher = make_watcher(tmp_path, observer, max_size_bytes=3)

    try:
        assert watcher.should_process(source) is False
        with qtbot.waitSignal(watcher.file_rejected, timeout=1_000) as signal:
            assert watcher.start() is True
        assert Path(signal.args[0]) == source
        assert "vượt giới hạn" in signal.args[1]
    finally:
        watcher.stop()


def test_update_settings_rebuilds_checker_and_restarts_observer(
    tmp_path: Path,
) -> None:
    old_output = tmp_path / "old"
    new_output = tmp_path / "new"
    observer = FakeObserver()
    watcher = make_watcher(old_output, observer)

    try:
        assert watcher.start() is True
        assert watcher.update_settings(
            {
                "output_dir": new_output,
            }
        )

        assert watcher.is_running is True
        assert watcher.output_dir == new_output
        assert watcher.file_patterns == ("ket_qua_boc_tach*.json",)
        assert watcher.stable_seconds == 0
        assert watcher.timeout_seconds == 1
        assert watcher.max_size_bytes == 1024
        assert observer.directory == str(new_output)
    finally:
        watcher.stop()
