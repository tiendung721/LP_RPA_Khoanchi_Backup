"""Kênh nội bộ để extension đóng đúng cửa sổ Trợ lý đã được mở."""

from __future__ import annotations

import json
import secrets
import time
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from threading import RLock, Thread
from urllib.parse import unquote, urlparse


@dataclass(frozen=True, slots=True)
class AssistantCloseSession:
    session_id: str
    port: int


@dataclass(slots=True)
class _SessionState:
    close_requested: bool
    created_at: float
    last_seen_at: float | None = None


class _BridgeServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True

    def __init__(self, bridge: "AssistantCloseBridge") -> None:
        self.bridge = bridge
        super().__init__(("127.0.0.1", 0), _BridgeHandler)


class _BridgeHandler(BaseHTTPRequestHandler):
    server: _BridgeServer

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        prefix = "/assistant-session/"
        if not parsed.path.startswith(prefix):
            self._send_json(404, {"close": False})
            return
        session_id = unquote(parsed.path[len(prefix) :])
        status = self.server.bridge.session_status(session_id)
        if status is None:
            self._send_json(404, {"close": False})
            return
        self._send_json(200, {"close": status})

    def _send_json(self, status: int, payload: dict[str, bool]) -> None:
        body = json.dumps(payload, separators=(",", ":")).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, _format: str, *_args: object) -> None:
        return


class AssistantCloseBridge:
    """Máy chủ loopback nhỏ; token ngẫu nhiên là quyền đóng một cửa sổ."""

    def __init__(self) -> None:
        self._lock = RLock()
        self._sessions: dict[str, _SessionState] = {}
        self._server: _BridgeServer | None = None
        self._thread: Thread | None = None

    @property
    def port(self) -> int | None:
        with self._lock:
            if self._server is None:
                return None
            return int(self._server.server_address[1])

    def create_session(self) -> AssistantCloseSession:
        with self._lock:
            self._ensure_started()
            session_id = secrets.token_urlsafe(32)
            self._sessions[session_id] = _SessionState(False, time.monotonic())
            assert self._server is not None
            return AssistantCloseSession(
                session_id=session_id,
                port=int(self._server.server_address[1]),
            )

    def request_close(self, session_id: str) -> bool:
        with self._lock:
            if session_id not in self._sessions:
                return False
            self._sessions[session_id].close_requested = True
            return True

    def discard_session(self, session_id: str) -> None:
        with self._lock:
            self._sessions.pop(session_id, None)

    def session_status(self, session_id: str) -> bool | None:
        with self._lock:
            state = self._sessions.get(session_id)
            if state is None:
                return None
            state.last_seen_at = time.monotonic()
            return state.close_requested

    def is_session_active(
        self,
        session_id: str,
        *,
        startup_grace_seconds: float = 15.0,
        heartbeat_timeout_seconds: float = 10.0,
    ) -> bool:
        with self._lock:
            state = self._sessions.get(session_id)
            if state is None:
                return False
            now = time.monotonic()
            if state.last_seen_at is None:
                return now - state.created_at <= startup_grace_seconds
            return now - state.last_seen_at <= heartbeat_timeout_seconds

    def close(self) -> None:
        with self._lock:
            server = self._server
            thread = self._thread
            self._server = None
            self._thread = None
            self._sessions.clear()
        if server is None:
            return
        server.shutdown()
        server.server_close()
        if thread is not None and thread.is_alive():
            thread.join(timeout=2)

    def _ensure_started(self) -> None:
        if self._server is not None:
            return
        server = _BridgeServer(self)
        thread = Thread(
            target=server.serve_forever,
            name="assistant-close-bridge",
            daemon=True,
        )
        self._server = server
        self._thread = thread
        thread.start()


__all__ = ["AssistantCloseBridge", "AssistantCloseSession"]
