"""Cấu hình ứng dụng và cây thư mục runtime."""

from __future__ import annotations

import filecmp
import json
import os
import shutil
import stat
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Mapping

from app.constants import DEFAULT_MAX_FILE_SIZE_BYTES


def software_root() -> Path:
    """Thư mục chứa source hoặc executable đang chạy."""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def default_data_root() -> Path:
    """Mặc định lưu dữ liệu cạnh source hoặc executable để bundle có tính di động."""

    return software_root()


def default_output_dir() -> Path:
    return software_root() / "Output"


@dataclass(frozen=True, slots=True)
class AppPaths:
    data_root: Path
    config_dir: Path
    settings_path: Path
    output_dir: Path
    system_dir: Path
    archive_original_dir: Path
    workspace_dir: Path
    ready_dir: Path
    rejected_dir: Path
    excel_dir: Path
    excel_temp_dir: Path
    excel_backup_dir: Path
    excel_reports_dir: Path
    rpa_dir: Path
    database_dir: Path
    database_path: Path
    logs_dir: Path
    log_path: Path

    @classmethod
    def from_data_root(
        cls,
        data_root: str | Path,
        output_dir: str | Path | None = None,
    ) -> "AppPaths":
        root = Path(data_root).expanduser()
        output = Path(output_dir).expanduser() if output_dir else default_output_dir()
        config_dir = root / "Config"
        database_dir = root / "Database"
        logs_dir = root / "Logs"
        system_dir = output / "_system"
        excel_dir = system_dir / "Excel"
        return cls(
            data_root=root,
            config_dir=config_dir,
            settings_path=config_dir / "settings.json",
            output_dir=output,
            system_dir=system_dir,
            archive_original_dir=system_dir / "Archive" / "Original",
            workspace_dir=system_dir / "Workspace",
            ready_dir=system_dir / "Ready",
            rejected_dir=system_dir / "Rejected",
            excel_dir=excel_dir,
            excel_temp_dir=excel_dir / "Temp",
            excel_backup_dir=excel_dir / "Backup",
            excel_reports_dir=excel_dir / "Reports",
            rpa_dir=system_dir / "RPA",
            database_dir=database_dir,
            database_path=database_dir / "app_state.db",
            logs_dir=logs_dir,
            log_path=logs_dir / "app.log",
        )

    @classmethod
    def defaults(cls) -> "AppPaths":
        return cls.from_data_root(default_data_root())

    def ensure_directories(self) -> None:
        for directory in self.directories:
            directory.mkdir(parents=True, exist_ok=True)

    @property
    def directories(self) -> tuple[Path, ...]:
        return (
            self.data_root,
            self.config_dir,
            self.output_dir,
            self.system_dir,
            self.excel_dir,
            self.excel_temp_dir,
            self.excel_backup_dir,
            self.excel_reports_dir,
            self.rpa_dir,
            self.database_dir,
            self.logs_dir,
        )

    @property
    def archive_dir(self) -> Path:
        return self.archive_original_dir

    @property
    def db_path(self) -> Path:
        return self.database_path

    @property
    def legacy_runtime_rebases(self) -> tuple[tuple[Path, Path], ...]:
        """Các cặp đường dẫn từ layout phẳng cũ sang ``Output/_system``."""

        return (
            (self.data_root / "Archive", self.system_dir / "Archive"),
            (self.data_root / "Workspace", self.workspace_dir),
            (self.data_root / "Ready", self.ready_dir),
            (self.data_root / "Rejected", self.rejected_dir),
        )


@dataclass(slots=True)
class AppSettings:
    data_root: Path = field(default_factory=default_data_root)
    assistant_bat_path: str = ""
    bang_ke_assistant_bat_path: str = ""
    output_dir: Path = field(default_factory=default_output_dir)
    daily_workbook_path: str = ""
    bk_workbook_path: str = ""
    payment_workbook_path: str = ""
    rpa_expense_bat_path: str = ""

    def __post_init__(self) -> None:
        if not isinstance(self.data_root, (str, os.PathLike)) or not str(
            self.data_root
        ).strip():
            raise ValueError("Thư mục dữ liệu không hợp lệ.")
        if not isinstance(self.output_dir, (str, os.PathLike)) or not str(
            self.output_dir
        ).strip():
            raise ValueError("Thư mục Output không được để trống.")
        optional_paths = (
            ("assistant_bat_path", "Đường dẫn BAT"),
            ("bang_ke_assistant_bat_path", "Đường dẫn BAT tool bảng kê"),
            ("daily_workbook_path", "Đường dẫn file Hàng ngày"),
            ("bk_workbook_path", "Đường dẫn file BK"),
            (
                "payment_workbook_path",
                "Đường dẫn file Thanh toán Nâng hạ/VS D/O",
            ),
            ("rpa_expense_bat_path", "Đường dẫn BAT RPA nhập quyết toán"),
        )
        for attribute, label in optional_paths:
            value = getattr(self, attribute)
            if isinstance(value, os.PathLike):
                value = os.fspath(value)
            if not isinstance(value, str):
                raise ValueError(f"{label} phải là chuỗi.")
            setattr(self, attribute, value.strip())
        self.data_root = Path(self.data_root).expanduser()
        self.output_dir = Path(self.output_dir).expanduser()

    @property
    def paths(self) -> AppPaths:
        return AppPaths.from_data_root(self.data_root, self.output_dir)

    @property
    def max_file_size_bytes(self) -> int:
        return DEFAULT_MAX_FILE_SIZE_BYTES

    def to_dict(self) -> dict[str, Any]:
        return {
            "data_root": str(self.data_root),
            "assistant_bat_path": self.assistant_bat_path,
            "bang_ke_assistant_bat_path": self.bang_ke_assistant_bat_path,
            "output_dir": str(self.output_dir),
            "daily_workbook_path": self.daily_workbook_path,
            "bk_workbook_path": self.bk_workbook_path,
            "payment_workbook_path": self.payment_workbook_path,
            "rpa_expense_bat_path": self.rpa_expense_bat_path,
        }

    @classmethod
    def from_dict(
        cls,
        value: Mapping[str, Any],
        *,
        fallback_paths: AppPaths | None = None,
    ) -> "AppSettings":
        fallback = fallback_paths or AppPaths.defaults()
        return cls(
            data_root=value.get("data_root", fallback.data_root),
            assistant_bat_path=value.get("assistant_bat_path", ""),
            bang_ke_assistant_bat_path=value.get("bang_ke_assistant_bat_path", ""),
            output_dir=value.get("output_dir", fallback.output_dir),
            daily_workbook_path=value.get("daily_workbook_path", ""),
            bk_workbook_path=value.get("bk_workbook_path", ""),
            payment_workbook_path=value.get("payment_workbook_path", ""),
            rpa_expense_bat_path=value.get("rpa_expense_bat_path", ""),
        )


Settings = AppSettings


class ConfigManager:
    """Nạp/lưu settings UTF-8 bằng phép thay thế file nguyên tử."""

    def __init__(
        self,
        data_root: str | Path | AppPaths | None = None,
        *,
        paths: AppPaths | None = None,
    ) -> None:
        if isinstance(data_root, AppPaths):
            if paths is not None:
                raise TypeError("Không truyền đồng thời data_root AppPaths và paths.")
            paths = data_root
        if paths is not None:
            self._bootstrap_paths = paths
        elif data_root is not None:
            self._bootstrap_paths = AppPaths.from_data_root(data_root)
        else:
            self._bootstrap_paths = AppPaths.defaults()
        self._relocated_from: Path | None = None

    @property
    def settings_path(self) -> Path:
        return self._bootstrap_paths.settings_path

    @property
    def relocated_from(self) -> Path | None:
        """Gốc dữ liệu ghi trong settings trước khi bundle được chuyển vị trí."""

        return self._relocated_from

    def load(self) -> AppSettings:
        self._relocated_from = None
        self._bootstrap_paths.ensure_directories()
        if not self.settings_path.exists():
            settings = AppSettings(
                data_root=self._bootstrap_paths.data_root,
                output_dir=self._bootstrap_paths.output_dir,
            )
            self.save(settings)
            return settings
        try:
            raw = json.loads(self.settings_path.read_text(encoding="utf-8-sig"))
        except (OSError, UnicodeError, json.JSONDecodeError) as exc:
            raise ValueError(
                f"Không thể đọc cấu hình tại {self.settings_path}."
            ) from exc
        if not isinstance(raw, dict):
            raise ValueError("File cấu hình phải chứa một JSON object.")
        effective = dict(raw)
        stored_root = _coerce_path(raw.get("data_root"))
        bootstrap_root = self._bootstrap_paths.data_root
        if stored_root is not None and not _same_path(stored_root, bootstrap_root):
            self._relocated_from = stored_root
            stored_output = _coerce_path(raw.get("output_dir"))
            if stored_output is not None:
                rebased_output = _rebase_child_path(
                    stored_output,
                    old_root=stored_root,
                    new_root=bootstrap_root,
                )
                if rebased_output is not None:
                    effective["output_dir"] = str(rebased_output)
        # Gốc đã dùng để tìm Config luôn là nguồn sự thật. Quy tắc này giúp
        # toàn bộ bundle tiếp tục chạy khi được copy sang thư mục hoặc máy khác.
        effective["data_root"] = str(bootstrap_root)
        settings = AppSettings.from_dict(
            effective,
            fallback_paths=self._bootstrap_paths,
        )
        settings.paths.ensure_directories()
        # Ghi lại ngay để loại khóa cấu hình cũ và cố định vị trí bundle mới.
        if raw != settings.to_dict():
            self.save(settings)
        return settings

    def save(self, settings: AppSettings) -> Path:
        target = self.settings_path
        target.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(
            settings.to_dict(),
            ensure_ascii=False,
            indent=2,
            allow_nan=False,
        ).encode("utf-8")
        temporary: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="wb",
                prefix=f".{target.name}.",
                suffix=".tmp",
                dir=target.parent,
                delete=False,
            ) as handle:
                temporary = Path(handle.name)
                handle.write(payload)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary, target)
            temporary = None
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
        settings.paths.ensure_directories()
        return target


def _coerce_path(value: object) -> Path | None:
    if not isinstance(value, (str, os.PathLike)) or not str(value).strip():
        return None
    return Path(value).expanduser()


def _same_path(left: Path, right: Path) -> bool:
    return os.path.normcase(str(left.resolve(strict=False))) == os.path.normcase(
        str(right.resolve(strict=False))
    )


def _rebase_child_path(
    path: Path,
    *,
    old_root: Path,
    new_root: Path,
) -> Path | None:
    try:
        relative = path.resolve(strict=False).relative_to(
            old_root.resolve(strict=False)
        )
    except ValueError:
        return None
    return new_root / relative


def migrate_legacy_runtime_layout(paths: AppPaths) -> int:
    """Gom dữ liệu của layout cũ vào ``Output/_system`` mà không ghi đè file."""

    moved_roots = 0
    for source, target in paths.legacy_runtime_rebases:
        if not source.exists():
            continue
        if not source.is_dir():
            raise ValueError(f"Đường dẫn dữ liệu cũ không phải thư mục: {source}")
        _merge_directory(source, target)
        moved_roots += 1
    return moved_roots


def _merge_directory(source: Path, target: Path) -> None:
    target.mkdir(parents=True, exist_ok=True)
    for child in tuple(source.iterdir()):
        destination = target / child.name
        if child.is_dir():
            if destination.exists() and not destination.is_dir():
                raise FileExistsError(
                    f"Không thể di trú vì đích đã là file: {destination}"
                )
            _merge_directory(child, destination)
            continue
        if destination.exists():
            if not destination.is_file() or not filecmp.cmp(
                child,
                destination,
                shallow=False,
            ):
                raise FileExistsError(
                    f"Không thể di trú vì file đích đã tồn tại: {destination}"
                )
            child.chmod(child.stat().st_mode | stat.S_IWRITE)
            child.unlink()
            continue
        shutil.move(str(child), str(destination))
    source.rmdir()


def load_settings(data_root: str | Path | None = None) -> AppSettings:
    return ConfigManager(data_root).load()


def save_settings(
    settings: AppSettings,
    data_root: str | Path | None = None,
) -> Path:
    return ConfigManager(data_root or settings.data_root).save(settings)


__all__ = [
    "AppPaths",
    "AppSettings",
    "ConfigManager",
    "Settings",
    "default_data_root",
    "default_output_dir",
    "load_settings",
    "migrate_legacy_runtime_layout",
    "save_settings",
    "software_root",
]
