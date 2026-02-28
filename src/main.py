"""audioflip — persistent, always-on-top audio device switcher widget.

Entry point for the application. Initialises all components and
starts the PyQt6 event loop.
"""

from __future__ import annotations

import sys
import logging

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QColor

from .config import ConfigManager
from .audio_manager import AudioManager
from .icons import IconManager
from .ui import AudioFlipWidget


def main() -> int:
    """Launch the audioflip widget."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

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
