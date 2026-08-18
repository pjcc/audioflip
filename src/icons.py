"""Icon management for audioflip.

Maps audio device names to appropriate icons using keyword matching,
with user-overridable per-device assignments stored in config.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from functools import lru_cache

from PyQt6.QtCore import QSize, Qt
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QColor
from PyQt6.QtSvg import QSvgRenderer


# All available icon types
ICON_TYPES: list[str] = [
    "headphones",
    "earbuds",
    "speaker",
    "bluetooth",
    "monitor",
    "monitor2",
    "tv",
    "usb",
    "microphone",
    "camera",
    "soundbar",
    "mixer",
    "phone",
    "audio",  # fallback
]

# Keyword-to-icon mapping, checked in order (first match wins)
_KEYWORD_MAP: list[tuple[list[str], str]] = [
    (["headphone", "headset"], "headphones"),
    (["earbud", "in-ear", "inear"], "earbuds"),
    (["bluetooth", "airpods", "buds", "galaxy buds"], "bluetooth"),
    (["monitor", "hdmi", "displayport", "display"], "monitor"),
    # After the display entry, so an HDMI-connected TV still reads as a monitor
    (["tv", "television"], "tv"),
    (["webcam", "camera"], "camera"),
    (["soundbar"], "soundbar"),
    (["dac", "audio interface", "interface"], "mixer"),
    (["phone", "hands-free", "handsfree"], "phone"),
    (["microphone", "mic"], "microphone"),
    (["usb"], "usb"),
    (["speaker", "realtek"], "speaker"),
]

FALLBACK_ICON = "audio"

# Everything get_icon will render: the device icons plus the menu checkmark.
_RENDERABLE_KEYS = frozenset(ICON_TYPES) | {"checkmark"}


def _resources_dir() -> Path:
    """Return path to the resources/ directory, handling frozen builds."""
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    else:
        base = Path(__file__).resolve().parent.parent
    return base / "resources"


def match_icon_for_name(device_name: str) -> str:
    """Return the best icon key for a device name using keyword matching.

    Args:
        device_name: The friendly name of the audio device.

    Returns:
        An icon key string (e.g. 'headphones', 'speaker').
    """
    lower = device_name.lower()
    for keywords, icon_key in _KEYWORD_MAP:
        for kw in keywords:
            if kw in lower:
                return icon_key
    return FALLBACK_ICON


class IconManager:
    """Provides QIcon instances for device icon keys, with tinting for dark theme."""

    def __init__(self, tint_color: QColor | None = None) -> None:
        self._tint = tint_color or QColor("#e0e0e0")
        self._resources = _resources_dir()
        self._cache: dict[tuple[str, int], QIcon] = {}

    def get_icon(self, icon_key: str, size: int = 20) -> QIcon:
        """Return a QIcon for the given icon key, tinted for dark theme.

        Args:
            icon_key: One of the ICON_TYPES keys.
            size: Pixel size for the icon.

        Returns:
            A tinted QIcon ready for display.
        """
        # icon_key can originate from icon_overrides in config.json, and it is
        # joined into a path below, so a value like '../../foo' would read an
        # SVG from outside resources/. Validate at the sink, which covers every
        # caller regardless of where the key came from.
        if icon_key not in _RENDERABLE_KEYS:
            icon_key = FALLBACK_ICON

        cache_key = (icon_key, size)
        if cache_key in self._cache:
            return self._cache[cache_key]

        svg_path = self._resources / f"{icon_key}.svg"
        if not svg_path.exists():
            svg_path = self._resources / f"{FALLBACK_ICON}.svg"

        icon = self._render_tinted_svg(svg_path, size)
        self._cache[cache_key] = icon
        return icon

    def get_checkmark_icon(self, size: int = 16) -> QIcon:
        """Return the checkmark icon, tinted."""
        return self.get_icon("checkmark", size)

    def _render_tinted_svg(self, svg_path: Path, size: int) -> QIcon:
        """Render an SVG file as a tinted QPixmap and wrap in QIcon."""
        renderer = QSvgRenderer(str(svg_path))
        pixmap = QPixmap(QSize(size, size))
        pixmap.fill(Qt.GlobalColor.transparent)

        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_SourceIn)
        painter.fillRect(pixmap.rect(), self._tint)
        painter.end()

        return QIcon(pixmap)

    def set_tint(self, color: QColor) -> None:
        """Change the tint color and clear the cache."""
        self._tint = color
        self._cache.clear()
