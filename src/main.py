"""audioflip — persistent, always-on-top audio device switcher widget.

Entry point for the application. Initialises all components and
starts the PyQt6 event loop.
"""

from __future__ import annotations

import os
import sys
import logging
from logging.handlers import RotatingFileHandler

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QColor

from .config import ConfigManager
from .audio_manager import AudioManager
from .icons import IconManager
from .ui import AudioFlipWidget


def _setup_logging() -> None:
    """Configure logging to stderr + rotating file in %APPDATA%/audioflip/."""
    fmt = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    level = logging.INFO

    # Always set up stderr (useful when running from terminal)
    logging.basicConfig(level=level, format=fmt)

    # Add a file handler so logs are available even with console=False
    log_dir = os.path.join(os.environ.get("APPDATA", "."), "audioflip")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, "audioflip.log")

    file_handler = RotatingFileHandler(
        log_path, maxBytes=1_000_000, backupCount=3, encoding="utf-8",
    )
    file_handler.setLevel(level)
    file_handler.setFormatter(logging.Formatter(fmt))
    logging.getLogger().addHandler(file_handler)


log = logging.getLogger(__name__)


def main() -> int:
    """Launch the audioflip widget."""
    _setup_logging()
    log.info("audioflip starting")

    app = QApplication(sys.argv)
    app.setApplicationName("audioflip")
    app.setQuitOnLastWindowClosed(False)

    config_mgr = ConfigManager.instance()
    audio_mgr = AudioManager()
    icon_mgr = IconManager(tint_color=QColor("#e0e0e0"))

    widget = AudioFlipWidget(audio_mgr, icon_mgr, config_mgr)
    widget.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
