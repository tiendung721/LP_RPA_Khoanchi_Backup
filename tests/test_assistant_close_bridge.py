from __future__ import annotations

import json
from urllib.error import HTTPError
from urllib.request import urlopen

import pytest

from app.services.assistant_close_bridge import AssistantCloseBridge


def _read(port: int, session_id: str) -> dict[str, bool]:
    with urlopen(
        f"http://127.0.0.1:{port}/assistant-session/{session_id}",
        timeout=2,
    ) as response:
        return json.loads(response.read())


def test_bridge_only_closes_the_requested_session() -> None:
    bridge = AssistantCloseBridge()
    first = bridge.create_session()
    second = bridge.create_session()
    try:
        assert _read(first.port, first.session_id) == {"close": False}
        assert _read(second.port, second.session_id) == {"close": False}

        assert bridge.request_close(first.session_id)
        assert _read(first.port, first.session_id) == {"close": True}
        assert _read(second.port, second.session_id) == {"close": False}
    finally:
        bridge.close()


def test_bridge_rejects_unknown_session() -> None:
    bridge = AssistantCloseBridge()
    session = bridge.create_session()
    try:
        with pytest.raises(HTTPError) as error:
            _read(session.port, "unknown-session-token-123456789")
        assert error.value.code == 404
    finally:
        bridge.close()
