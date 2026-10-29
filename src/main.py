"""audioflip — persistent, always-on-top audio device switcher widget.

Entry point for the application. Initialises all components and
starts the PyQt6 event loop.
"""

from __future__ import annotations

import datetime
import faulthandler
import getpass
import os
import sys
import logging
from logging.handlers import RotatingFileHandler

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QColor
from PyQt6.QtNetwork import QLocalServer, QLocalSocket

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


_crash_file = None  # module-level so the handle outlives _setup_crash_handler


def _setup_crash_handler() -> None:
    """Write a Python traceback to disk if the process dies in native code.

    The packaged build runs with console=False, so stderr goes nowhere. A
    native crash - an access violation inside a ctypes or COM call, say -
    then kills the process silently, leaving only a log that stops
    mid-sentence. faulthandler dumps the stack of every thread instead,
    which is the difference between a diagnosable crash and a mystery.
    """
    global _crash_file
    try:
        log_dir = os.path.join(os.environ.get("APPDATA", "."), "audioflip")
        os.makedirs(log_dir, exist_ok=True)
        _crash_file = open(
            os.path.join(log_dir, "crash.log"), "a", buffering=1, encoding="utf-8",
        )
        _crash_file.write(
            "\n===== session started "
            + datetime.datetime.now().isoformat(timespec="seconds")
            + " =====\n"
        )
        faulthandler.enable(file=_crash_file, all_threads=True)
    except Exception:
        pass  # diagnostics must never stop the app starting


log = logging.getLogger(__name__)

# Named pipe the running instance listens on. Per user, since pipe names
# are machine-wide.
_INSTANCE_KEY = f"audioflip-{getpass.getuser()}"


def _describe_build() -> str:
    """Name the binary that is starting, for the log.

    Several copies of audioflip.exe have existed on one machine at once,
    and an old one launched from search looked like the app had lost
    features. With this line the log says which copy ran.
    """
    if not getattr(sys, "frozen", False):
        return "from source"
    try:
        built = datetime.datetime.fromtimestamp(os.path.getmtime(sys.executable))
        return f"{sys.executable}, built {built:%Y-%m-%d %H:%M}"
    except OSError:
        return sys.executable


def _hand_off_to_running_instance() -> bool:
    """Ask an already-running audioflip to surface itself.

    Returns True if one answered, in which case this process should exit
    rather than start a second widget with its own timers and tray icon.
    """
    socket = QLocalSocket()
    socket.connectToServer(_INSTANCE_KEY)
    if not socket.waitForConnected(500):
        return False
    socket.disconnectFromServer()
    return True


def _listen_for_later_launches(widget: AudioFlipWidget) -> QLocalServer:
    """Surface *widget* whenever another launch hands off to us.

    Parented to the widget, so Qt keeps the listener alive for as long as
    the widget exists.
    """
    server = QLocalServer(widget)
    if not server.listen(_INSTANCE_KEY):
        log.warning("Single-instance listener unavailable: %s", server.errorString())

    def _on_connection() -> None:
        conn = server.nextPendingConnection()
        if conn is not None:
            conn.disconnectFromServer()
            conn.deleteLater()
        log.info("Another launch handed off - surfacing the widget")
        widget.surface()

    server.newConnection.connect(_on_connection)
    return server


def main() -> int:
    """Launch the audioflip widget."""
    _setup_logging()

    app = QApplication(sys.argv)
    app.setApplicationName("audioflip")
    app.setQuitOnLastWindowClosed(False)

    if _hand_off_to_running_instance():
        log.info("audioflip already running (%s) - surfaced it instead", _describe_build())
        return 0
    _setup_crash_handler()
    log.info("audioflip starting (%s)", _describe_build())

    config_mgr = ConfigManager.instance()
    audio_mgr = AudioManager()
    icon_mgr = IconManager(tint_color=QColor("#e0e0e0"))

    widget = AudioFlipWidget(audio_mgr, icon_mgr, config_mgr)
    _listen_for_later_launches(widget)
    widget.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
