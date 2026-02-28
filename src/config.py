"""Configuration management for audioflip.

Handles loading, saving, and accessing application settings
from %APPDATA%/audioflip/config.json.
"""

from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any


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
    flash_on_change: bool = True

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
                flash_on_change=data.get("flash_on_change", True),
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

    def toggle_favourite(self, device_id: str) -> bool:
        """Toggle a device as favourite. Returns True if now favourite."""
        if device_id in self._config.favourites:
            self._config.favourites.remove(device_id)
            self.save()
            return False
        self._config.favourites.append(device_id)
        self.save()
        return True

    def is_favourite(self, device_id: str) -> bool:
        return device_id in self._config.favourites

    def set_flash_on_change(self, value: bool) -> None:
        self._config.flash_on_change = value
        self.save()

    @staticmethod
    def _update_startup_shortcut(enable: bool) -> None:
        """Add or remove from Windows startup via the Start Menu Startup folder."""
        try:
            startup_dir = Path(
                os.environ.get(
                    "APPDATA",
                    Path.home() / "AppData" / "Roaming",
                )
            ) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
            shortcut_path = startup_dir / "audioflip.lnk"

            if not enable:
                if shortcut_path.exists():
                    shortcut_path.unlink()
                return

            # Determine the executable path
            if getattr(sys, "frozen", False):
                target = sys.executable
            else:
                target = str(Path(sys.executable))
                # When running from source, we'd need a different approach;
                # for the packaged .exe this will be correct.

            # Create shortcut using PowerShell
            ps_script = (
                f'$ws = New-Object -ComObject WScript.Shell; '
                f'$s = $ws.CreateShortcut("{shortcut_path}"); '
                f'$s.TargetPath = "{target}"; '
                f'$s.WorkingDirectory = "{Path(target).parent}"; '
                f'$s.Save()'
            )
            import subprocess
            subprocess.run(
                ["powershell", "-Command", ps_script],
                capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        except Exception:
            pass  # Non-critical; don't crash the app
