"""Configuration management for audioflip.

Handles loading, saving, and accessing application settings
from %APPDATA%/audioflip/config.json.
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any

log = logging.getLogger(__name__)


def _ps_quote(value: str) -> str:
    """Quote a value for use as a PowerShell single-quoted literal."""
    return "'" + str(value).replace("'", "''") + "'"


def _appdata_dir() -> Path:
    """Return the application data directory, creating it if needed."""
    base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    app_dir = base / "audioflip"
    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir


VALID_THEMES = (
    "dark", "light", "midnight",
    "ocean", "forest", "sunset", "berry", "slate", "copper", "arctic",
)


@dataclass
class Config:
    """Application configuration with sensible defaults."""

    always_on_top: bool = True
    position: dict[str, int] = field(default_factory=lambda: {"x": 100, "y": 100})
    show_mode: str = "output"  # "output" | "input" | "both"
    start_with_windows: bool = False
    icon_overrides: dict[str, str] = field(default_factory=dict)
    theme: str = "dark"
    favourites: list[str] = field(default_factory=list)  # device IDs pinned to top
    favourite_devices: dict[str, dict] = field(default_factory=dict)  # {id: {name, flow, is_bluetooth}}
    flash_on_change: bool = True
    show_volume_bar: bool = False
    yield_to_fullscreen: bool = True  # drop topmost while a fullscreen app is active

    def __post_init__(self) -> None:
        if self.show_mode not in ("output", "input", "both"):
            self.show_mode = "output"
        if self.theme not in VALID_THEMES:
            self.theme = "dark"


class ConfigManager:
    """Singleton-style manager for loading/saving config.json."""

    _instance: ConfigManager | None = None
    _config: Config
    _path: Path

    def __init__(self, config_path: Path | None = None) -> None:
        self._path = config_path or (_appdata_dir() / "config.json")
        self._config = self._load()
        self._reconcile_startup_shortcut()

    @classmethod
    def instance(cls, config_path: Path | None = None) -> ConfigManager:
        """Return the shared ConfigManager, creating it on first call."""
        if cls._instance is None:
            cls._instance = cls(config_path)
        return cls._instance

    @classmethod
    def reset(cls) -> None:
        """Reset the singleton (useful for tests)."""
        cls._instance = None

    @property
    def config(self) -> Config:
        return self._config

    @property
    def path(self) -> Path:
        return self._path

    def _load(self) -> Config:
        """Load config from disk, returning defaults on any error."""
        if not self._path.exists():
            cfg = Config()
            self._save(cfg)
            return cfg
        try:
            data: dict[str, Any] = json.loads(self._path.read_text(encoding="utf-8"))
            return Config(
                always_on_top=data.get("always_on_top", True),
                position=data.get("position", {"x": 100, "y": 100}),
                show_mode=data.get("show_mode", "output"),
                start_with_windows=data.get("start_with_windows", False),
                icon_overrides=data.get("icon_overrides", {}),
                theme=data.get("theme", "dark"),
                favourites=data.get("favourites", []),
                favourite_devices=data.get("favourite_devices", {}),
                flash_on_change=data.get("flash_on_change", True),
                show_volume_bar=data.get("show_volume_bar", False),
                yield_to_fullscreen=data.get("yield_to_fullscreen", True),
            )
        except (json.JSONDecodeError, KeyError, TypeError):
            return Config()

    def _save(self, cfg: Config) -> None:
        """Persist config to disk."""
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self._path.write_text(
            json.dumps(asdict(cfg), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    def save(self) -> None:
        """Save current config to disk."""
        self._save(self._config)

    def set_position(self, x: int, y: int) -> None:
        self._config.position = {"x": x, "y": y}
        self.save()

    def set_always_on_top(self, value: bool) -> None:
        self._config.always_on_top = value
        self.save()

    def set_show_mode(self, mode: str) -> None:
        if mode in ("output", "input", "both"):
            self._config.show_mode = mode
            self.save()

    def set_start_with_windows(self, value: bool) -> None:
        self._config.start_with_windows = value
        self._update_startup_shortcut(value)
        self.save()

    def set_icon_override(self, device_id: str, icon_name: str) -> None:
        self._config.icon_overrides[device_id] = icon_name
        self.save()

    def get_icon_override(self, device_id: str) -> str | None:
        return self._config.icon_overrides.get(device_id)

    def set_theme(self, theme: str) -> None:
        if theme in VALID_THEMES:
            self._config.theme = theme
            self.save()

    def toggle_favourite(
        self,
        device_id: str,
        device_name: str = "",
        flow: str = "output",
        is_bluetooth: bool = False,
    ) -> bool:
        """Toggle a device as favourite. Returns True if now favourite.

        When adding, stores metadata so disconnected devices can appear as
        ghost entries in the dropdown.
        """
        if device_id in self._config.favourites:
            self._config.favourites.remove(device_id)
            self._config.favourite_devices.pop(device_id, None)
            self.save()
            return False
        self._config.favourites.append(device_id)
        self._config.favourite_devices[device_id] = {
            "name": device_name,
            "flow": flow,
            "is_bluetooth": is_bluetooth,
        }
        self.save()
        return True

    def is_favourite(self, device_id: str) -> bool:
        return device_id in self._config.favourites

    def get_favourite_devices(self) -> dict[str, dict]:
        """Return the favourite device metadata dict."""
        return self._config.favourite_devices

    def migrate_favourite_id(self, old_id: str, new_id: str, new_name: str = "") -> None:
        """Replace *old_id* with *new_id* in favourites, metadata, and icon overrides.

        Used when a Bluetooth device reconnects with a different endpoint ID
        but the same underlying device name.
        """
        if old_id not in self._config.favourites:
            return

        # Swap in favourites list (preserve order)
        idx = self._config.favourites.index(old_id)
        self._config.favourites[idx] = new_id

        # Migrate metadata
        meta = self._config.favourite_devices.pop(old_id, None)
        if meta:
            if new_name:
                meta["name"] = new_name
            self._config.favourite_devices[new_id] = meta

        # Migrate icon override
        icon = self._config.icon_overrides.pop(old_id, None)
        if icon is not None:
            self._config.icon_overrides[new_id] = icon

        self.save()

    def set_flash_on_change(self, value: bool) -> None:
        self._config.flash_on_change = value
        self.save()

    def set_show_volume_bar(self, value: bool) -> None:
        self._config.show_volume_bar = value
        self.save()

    def set_yield_to_fullscreen(self, value: bool) -> None:
        self._config.yield_to_fullscreen = value
        self.save()

    @staticmethod
    def _startup_shortcut_path() -> Path:
        """Return the path of the Start Menu Startup shortcut."""
        base = Path(
            os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")
        )
        return (
            base / "Microsoft" / "Windows" / "Start Menu" / "Programs"
            / "Startup" / "audioflip.lnk"
        )

    @staticmethod
    def _startup_target() -> tuple[str, str, str]:
        """Return (target, arguments, working_dir) for the startup shortcut.

        Running from source needs the interpreter plus the script path; the
        old code pointed the shortcut at a bare python.exe with no arguments,
        which launched an interpreter that did nothing. pythonw.exe is
        preferred so no console window appears at login.
        """
        if getattr(sys, "frozen", False):
            exe = Path(sys.executable)
            return str(exe), "", str(exe.parent)

        repo_root = Path(__file__).resolve().parent.parent
        script = repo_root / "run.py"
        interpreter = Path(sys.executable)
        windowless = interpreter.with_name("pythonw.exe")
        if not windowless.exists():
            windowless = interpreter
        return str(windowless), f'"{script}"', str(repo_root)

    def _reconcile_startup_shortcut(self) -> None:
        """Make the config agree with the Startup folder.

        The shortcut is what actually makes Windows launch audioflip, so it
        wins when the two disagree - which happens whenever config.json is
        recreated, or the shortcut is added or removed by hand. Without this
        the menu shows a checkbox that contradicts what the machine does.
        """
        actual = self._startup_shortcut_path().exists()
        if actual != self._config.start_with_windows:
            log.info(
                "Startup shortcut %s but config said %s - trusting the shortcut",
                "exists" if actual else "is absent",
                self._config.start_with_windows,
            )
            self._config.start_with_windows = actual
            self.save()

    @classmethod
    def _update_startup_shortcut(cls, enable: bool) -> None:
        """Add or remove the Start Menu Startup shortcut."""
        try:
            shortcut_path = cls._startup_shortcut_path()

            if not enable:
                if shortcut_path.exists():
                    shortcut_path.unlink()
                return

            target, arguments, workdir = cls._startup_target()
            shortcut_path.parent.mkdir(parents=True, exist_ok=True)

            ps_script = (
                "$ws = New-Object -ComObject WScript.Shell; "
                f"$s = $ws.CreateShortcut({_ps_quote(str(shortcut_path))}); "
                f"$s.TargetPath = {_ps_quote(target)}; "
                f"$s.Arguments = {_ps_quote(arguments)}; "
                f"$s.WorkingDirectory = {_ps_quote(workdir)}; "
                "$s.Save()"
            )
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_script],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            if result.returncode != 0:
                log.warning(
                    "Failed to create startup shortcut: %s",
                    (result.stderr or "").strip(),
                )
        except Exception as exc:
            log.warning("Startup shortcut update failed: %s", exc)
