from __future__ import annotations

import json
import os
import time
from pathlib import Path

import pytest

import app.services.assistant_bat_launcher as launcher_module
from app.services.assistant_bat_launcher import (
    AssistantBatLauncher,
    AssistantLaunchError,
)


def _bundle(tmp_path: Path) -> tuple[Path, Path, Path]:
    root = tmp_path / "assistant"
    launcher_dir = root / "RPA_ChatGPT_Launcher"
    extension_dir = root / "RPA_ChatGPT_Extension"
    profile_dir = root / "RPA_ChatGPT_Profile"
    launcher_dir.mkdir(parents=True)
    extension_dir.mkdir()
    (extension_dir / "manifest.json").write_text(
        json.dumps({"manifest_version": 3, "version": "2.2.0"}),
        encoding="utf-8",
    )
    (extension_dir / "window-manager.js").write_text("", encoding="utf-8")
    (extension_dir / "title-cleaner.js").write_text("", encoding="utf-8")
    bat = launcher_dir / "Mo_Tro_Ly_RPA.bat"
    bat.write_text("@echo off\r\n", encoding="utf-8")
    (launcher_dir / "open_chatgpt_app.ps1").write_text(
        "# launcher", encoding="utf-8"
    )
    return bat, profile_dir, root


def test_preferences_merge_preserves_existing_profile_values(tmp_path: Path) -> None:
    bat, profile_dir, _ = _bundle(tmp_path)
    preferences = profile_dir / "Default" / "Preferences"
    preferences.parent.mkdir(parents=True)
    preferences.write_text(
        json.dumps(
            {
                "account_info": [{"email": "user@example.test"}],
                "download": {"old_key": "kept"},
            }
        ),
        encoding="utf-8",
    )
    output = tmp_path / "Output"
    launcher = AssistantBatLauncher()

    bundle = launcher.validate_configuration(bat, output)
    launcher.configure_download_directory(bundle, output)
    saved = json.loads(preferences.read_text(encoding="utf-8"))

    assert saved["account_info"] == [{"email": "user@example.test"}]
    assert saved["download"]["old_key"] == "kept"
    assert saved["download"]["default_directory"] == str(output.resolve())
    assert saved["download"]["prompt_for_download"] is False
    assert saved["download"]["directory_upgrade"] is True


def test_launch_runs_bat_detached(tmp_path: Path, monkeypatch) -> None:
    bat, _, _ = _bundle(tmp_path)
    output = tmp_path / "Output"
    calls: list[tuple[str, list[str], str]] = []

    def start_detached(
        program: str, arguments: list[str], working_directory: str
    ) -> tuple[bool, int]:
        calls.append((program, arguments, working_directory))
        return True, 321

    monkeypatch.setattr(launcher_module.QProcess, "startDetached", start_detached)

    launcher = AssistantBatLauncher()
    result = launcher.launch(bat, output)

    assert result.success
    assert result.process_id == 321
    assert calls[0][1][:3] == ["/d", "/s", "/c"]
    assert calls[0][1][3] == "call"
    assert calls[0][1][4] == str(bat.resolve())
    assert len(calls[0][1][5]) >= 20
    assert int(calls[0][1][6]) > 0
    assert calls[0][2] == str(bat.parent.resolve())
    assert result.session_id == calls[0][1][5]
    assert result.bridge_port == int(calls[0][1][6])
    launcher.close()


@pytest.mark.skipif(os.name != "nt", reason="Luồng BAT chỉ hỗ trợ Windows.")
def test_detached_command_really_executes_bat_with_spaces(
    tmp_path: Path,
) -> None:
    root = tmp_path / "bundle có khoảng trắng"
    launcher_dir = root / "RPA_ChatGPT_Launcher"
    extension_dir = root / "RPA_ChatGPT_Extension"
    launcher_dir.mkdir(parents=True)
    extension_dir.mkdir()
    (extension_dir / "manifest.json").write_text(
        json.dumps({"manifest_version": 3, "version": "2.2.0"}),
        encoding="utf-8",
    )
    (extension_dir / "window-manager.js").write_text("", encoding="utf-8")
    (extension_dir / "title-cleaner.js").write_text("", encoding="utf-8")
    marker = launcher_dir / "started.ok"
    bat = launcher_dir / "Mở Trợ Lý.bat"
    bat.write_text(
        '@echo off\r\n>"%~dp0started.ok" echo started\r\n',
        encoding="utf-8",
    )
    (launcher_dir / "open_chatgpt_app.ps1").write_text(
        "# launcher",
        encoding="utf-8",
    )

    result = AssistantBatLauncher().launch(bat, root / "Output")
    deadline = time.monotonic() + 3
    while time.monotonic() < deadline and not marker.is_file():
        time.sleep(0.05)

    assert result.success
    assert marker.is_file()


@pytest.mark.parametrize(
    "bat_name",
    ["", "missing.bat", "assistant.exe"],
)
def test_invalid_bat_is_rejected(tmp_path: Path, bat_name: str) -> None:
    launcher = AssistantBatLauncher()
    with pytest.raises(AssistantLaunchError):
        launcher.validate_configuration(
            tmp_path / bat_name if bat_name else "",
            tmp_path / "Output",
        )


def test_explicit_empty_bat_does_not_fall_back_to_saved_setting(
    tmp_path: Path,
) -> None:
    bat, _, _ = _bundle(tmp_path)
    launcher = AssistantBatLauncher(
        {
            "assistant_bat_path": str(bat),
            "output_dir": str(tmp_path / "Output"),
        }
    )

    with pytest.raises(AssistantLaunchError, match="Chưa cấu hình"):
        launcher.validate_configuration("", tmp_path / "Output")


def test_explicit_empty_output_does_not_use_working_directory(
    tmp_path: Path,
) -> None:
    bat, _, _ = _bundle(tmp_path)
    launcher = AssistantBatLauncher(
        {
            "assistant_bat_path": str(bat),
            "output_dir": str(tmp_path / "Output"),
        }
    )

    with pytest.raises(AssistantLaunchError, match="Chưa cấu hình"):
        launcher.launch(bat, "")


def test_old_extension_bundle_is_rejected(tmp_path: Path) -> None:
    bat, _, root = _bundle(tmp_path)
    manifest = root / "RPA_ChatGPT_Extension" / "manifest.json"
    manifest.write_text(
        json.dumps({"manifest_version": 3, "version": "2.1.0"}),
        encoding="utf-8",
    )

    with pytest.raises(AssistantLaunchError, match="2.2.0"):
        AssistantBatLauncher().validate_configuration(bat, tmp_path / "Output")
