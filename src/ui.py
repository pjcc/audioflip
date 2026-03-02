"""UI components for the audioflip widget.

Contains the main widget (compact icon + name), the device dropdown,
the right-click context menu, and the system tray icon.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import logging
import time
from typing import TYPE_CHECKING

from PyQt6.QtCore import (
    Qt,
    QEvent,
    QObject,
    QPoint,
    QPropertyAnimation,
    QEasingCurve,
    QRectF,
    QThread,
    QTimer,
    pyqtSignal,
    QSize,
    QRect,
)
from PyQt6.QtGui import (
    QAction,
    QColor,
    QCursor,
    QFont,
    QIcon,
    QMouseEvent,
    QPainter,
    QPainterPath,
    QPaintEvent,
    QPen,
    QPixmap,
    QWheelEvent,
)
from PyQt6.QtWidgets import (
    QApplication,
    QGraphicsOpacityEffect,
    QHBoxLayout,
    QLabel,
    QMenu,
    QScrollArea,
    QSizePolicy,
    QStyle,
    QStyleOption,
    QSystemTrayIcon,
    QVBoxLayout,
    QWidget,
)

from .audio_manager import AudioDevice, AudioManager, DeviceFlow
from .bluetooth import bluetooth_connect, bluetooth_disconnect, is_bluetooth_available
from .config import ConfigManager, VALID_THEMES
from .icons import ICON_TYPES, IconManager, match_icon_for_name

if TYPE_CHECKING:
    pass

log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# BT name-matching helpers
# ---------------------------------------------------------------------------
def _bt_name_core(audio_name: str) -> str:
    """Extract the Bluetooth device name from a Windows audio endpoint name.

    Windows names BT audio endpoints like "Headphones (Buds Pro 2)".
    The actual BT device name is the part inside the last pair of parentheses.
    Returns the extracted name lower-cased, or the full name lower-cased
    if no parentheses are found.
    """
    close = audio_name.rfind(")")
    if close == -1:
        return audio_name.strip().lower()
    open_ = audio_name.rfind("(", 0, close)
    if open_ == -1:
        return audio_name.strip().lower()
    return audio_name[open_ + 1:close].strip().lower()


def _bt_names_match(name_a: str, name_b: str) -> bool:
    """Return True if two audio endpoint names refer to the same BT device.

    Compares the core BT device name extracted from each endpoint name.
    E.g. "Earphones (Buds Pro 2)" and "Headphones (Buds Pro 2)" both match.
    """
    core_a = _bt_name_core(name_a)
    core_b = _bt_name_core(name_b)
    return core_a == core_b and len(core_a) > 0


# ---------------------------------------------------------------------------
# Win32 constants for always-on-top above taskbar
# ---------------------------------------------------------------------------
_user32 = ctypes.windll.user32
_SWP_NOMOVE = 0x0002
_SWP_NOSIZE = 0x0001
_SWP_NOACTIVATE = 0x0010
_SWP_SHOWWINDOW = 0x0040
_HWND_TOPMOST = ctypes.wintypes.HWND(-1)
_HWND_NOTOPMOST = ctypes.wintypes.HWND(-2)

# ---------------------------------------------------------------------------
# Theme definitions
# ---------------------------------------------------------------------------
_THEMES: dict[str, dict[str, str]] = {
    "dark": {
        "bg": "#1e1e2e",
        "bg_hover": "#313244",
        "fg": "#cdd6f4",
        "fg_dim": "#a6adc8",
        "accent": "#89b4fa",
        "border": "#585b70",
        "icon_tint": "#e0e0e0",
    },
    "light": {
        "bg": "#eff1f5",
        "bg_hover": "#dce0e8",
        "fg": "#4c4f69",
        "fg_dim": "#6c6f85",
        "accent": "#1e66f5",
        "border": "#bcc0cc",
        "icon_tint": "#4c4f69",
    },
    "midnight": {
        "bg": "#0d1117",
        "bg_hover": "#161b22",
        "fg": "#c9d1d9",
        "fg_dim": "#8b949e",
        "accent": "#58a6ff",
        "border": "#30363d",
        "icon_tint": "#c9d1d9",
    },
    "ocean": {
        "bg": "#0f2942",
        "bg_hover": "#163557",
        "fg": "#d0e4f7",
        "fg_dim": "#7da8cc",
        "accent": "#4fc3f7",
        "border": "#1e4976",
        "icon_tint": "#d0e4f7",
    },
    "forest": {
        "bg": "#1a2e1a",
        "bg_hover": "#253d25",
        "fg": "#c8e6c9",
        "fg_dim": "#81a784",
        "accent": "#66bb6a",
        "border": "#2e5a2e",
        "icon_tint": "#c8e6c9",
    },
    "sunset": {
        "bg": "#2d1b2e",
        "bg_hover": "#3d2640",
        "fg": "#f0d6e8",
        "fg_dim": "#b08da3",
        "accent": "#f48fb1",
        "border": "#5a3555",
        "icon_tint": "#f0d6e8",
    },
    "berry": {
        "bg": "#1e1233",
        "bg_hover": "#2a1d47",
        "fg": "#e0d4f5",
        "fg_dim": "#9b8bb8",
        "accent": "#b388ff",
        "border": "#3d2b66",
        "icon_tint": "#e0d4f5",
    },
    "slate": {
        "bg": "#1e2830",
        "bg_hover": "#2a3740",
        "fg": "#c8d3db",
        "fg_dim": "#8899a6",
        "accent": "#78909c",
        "border": "#3d4f5f",
        "icon_tint": "#c8d3db",
    },
    "copper": {
        "bg": "#2a1f1a",
        "bg_hover": "#3a2c24",
        "fg": "#e8d5c4",
        "fg_dim": "#ad8e78",
        "accent": "#e6a070",
        "border": "#5a4335",
        "icon_tint": "#e8d5c4",
    },
    "arctic": {
        "bg": "#e8eef4",
        "bg_hover": "#d4dde8",
        "fg": "#2e3d4f",
        "fg_dim": "#5a7088",
        "accent": "#2196f3",
        "border": "#b8c8d8",
        "icon_tint": "#2e3d4f",
    },
}

_RADIUS = 10
_FONT_FAMILY = "Segoe UI"


def _t(config_mgr: ConfigManager) -> dict[str, str]:
    """Return the active theme palette dict."""
    return _THEMES.get(config_mgr.config.theme, _THEMES["dark"])


def _widget_stylesheet(t: dict[str, str]) -> str:
    return (
        f"QWidget#WidgetBody {{"
        f"  background-color: {t['bg']};"
        f"  border: 1px solid {t['border']};"
        f"  border-radius: {_RADIUS}px;"
        f"}}"
    )


def _dropdown_stylesheet(t: dict[str, str]) -> str:
    return (
        f"QWidget#Dropdown {{"
        f"  background-color: {t['bg']};"
        f"  border: 1px solid {t['border']};"
        f"  border-radius: {_RADIUS}px;"
        f"}}"
        f"QLabel#SectionHeader {{"
        f"  color: {t['fg_dim']};"
        f"  font-size: 11px;"
        f"  padding: 6px 12px 2px 12px;"
        f"}}"
    )


def _menu_stylesheet(t: dict[str, str]) -> str:
    return (
        f"QMenu {{ background-color: {t['bg']}; color: {t['fg']}; "
        f"border: 1px solid {t['border']}; border-radius: 6px; padding: 4px; }}"
        f"QMenu::item {{ padding: 6px 20px; border-radius: 4px; }}"
        f"QMenu::item:selected {{ background-color: {t['bg_hover']}; }}"
        f"QMenu::separator {{ background-color: {t['border']}; height: 1px; margin: 4px 8px; }}"
    )


def _scroll_stylesheet(t: dict[str, str]) -> str:
    return (
        f"QScrollArea {{ background: transparent; border: none; }}"
        f"QScrollBar:vertical {{ background: {t['bg']}; width: 6px; border-radius: 3px; }}"
        f"QScrollBar::handle:vertical {{ background: {t['border']}; border-radius: 3px; min-height: 20px; }}"
        f"QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}"
    )


# ---------------------------------------------------------------------------
# DeviceRow — single row in the dropdown
# ---------------------------------------------------------------------------
class DeviceRow(QWidget):
    """A single clickable device row in the dropdown."""

    clicked = pyqtSignal(object)  # emits AudioDevice
    fav_toggled = pyqtSignal(object)  # emits AudioDevice

    def __init__(
        self,
        device: AudioDevice,
        icon_mgr: IconManager,
        config_mgr: ConfigManager,
        is_fav: bool = False,
        disconnected: bool = False,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.device = device
        self._hovered = False
        self._is_fav = is_fav
        self._disconnected = disconnected
        self._theme = _t(config_mgr)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedHeight(36)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 4, 12, 4)
        layout.setSpacing(8)

        # Star for favourite — dim colour when disconnected
        self._star = QLabel("\u2605" if is_fav else "")
        self._star.setFixedWidth(14)
        star_colour = self._theme["fg_dim"] if disconnected else self._theme["accent"]
        self._star.setStyleSheet(
            f"color: {star_colour}; font-size: 12px; background: transparent;"
        )
        layout.addWidget(self._star)

        # Icon — draw at 40% opacity when disconnected
        override = config_mgr.get_icon_override(device.id)
        icon_key = override or match_icon_for_name(device.name)
        icon_label = QLabel()
        src_pixmap = icon_mgr.get_icon(icon_key, 18).pixmap(QSize(18, 18))
        if disconnected:
            faded = QPixmap(src_pixmap.size())
            faded.fill(Qt.GlobalColor.transparent)
            p = QPainter(faded)
            p.setOpacity(0.4)
            p.drawPixmap(0, 0, src_pixmap)
            p.end()
            icon_label.setPixmap(faded)
        else:
            icon_label.setPixmap(src_pixmap)
        icon_label.setFixedSize(18, 18)
        layout.addWidget(icon_label)

        # Name — dim colour when disconnected
        self._name_label = QLabel(device.name)
        self._name_label.setFont(QFont(_FONT_FAMILY, 10))
        name_colour = self._theme["fg_dim"] if disconnected else self._theme["fg"]
        self._name_label.setStyleSheet(
            f"color: {name_colour}; background: transparent;"
        )
        self._name_label.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        layout.addWidget(self._name_label)

        # Checkmark if default (never shown for disconnected)
        if device.is_default and not disconnected:
            check_label = QLabel()
            check_label.setPixmap(
                icon_mgr.get_checkmark_icon(14).pixmap(QSize(14, 14))
            )
            check_label.setFixedSize(14, 14)
            layout.addWidget(check_label)

    def set_name_text(self, text: str) -> None:
        """Update the displayed name (e.g. 'Connecting...')."""
        self._name_label.setText(text)

    def enterEvent(self, event: object) -> None:
        self._hovered = True
        self.update()

    def leaveEvent(self, event: object) -> None:
        self._hovered = False
        self.update()

    def paintEvent(self, event: QPaintEvent) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        if self._hovered:
            painter.setBrush(QColor(self._theme["bg_hover"]))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(self.rect(), 6, 6)
        painter.end()

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self.device)
        elif event.button() == Qt.MouseButton.RightButton:
            self.fav_toggled.emit(self.device)


# ---------------------------------------------------------------------------
# DeviceDropdown — popup list of devices
# ---------------------------------------------------------------------------
class DeviceDropdown(QWidget):
    """Popup dropdown showing all audio devices grouped by flow."""

    device_selected = pyqtSignal(object)  # emits AudioDevice
    bt_connect_requested = pyqtSignal(object)  # emits AudioDevice — BT connect
    bt_disconnect_requested = pyqtSignal(object)  # emits AudioDevice — BT disconnect
    favourite_toggled = pyqtSignal(object)  # emits AudioDevice
    closed = pyqtSignal()  # emitted when the dropdown hides

    def __init__(
        self,
        audio_mgr: AudioManager,
        icon_mgr: IconManager,
        config_mgr: ConfigManager,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(
            parent, Qt.WindowType.Popup | Qt.WindowType.FramelessWindowHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setObjectName("Dropdown")
        self._audio_mgr = audio_mgr
        self._icon_mgr = icon_mgr
        self._config_mgr = config_mgr
        self._rows_by_device_id: dict[str, DeviceRow] = {}
        self._bt_busy = False  # True while a BT op is in progress
        self._last_show_mode = "output"
        self._last_widget_height = 40

        # Opacity animation
        self._opacity = QGraphicsOpacityEffect(self)
        self._opacity.setOpacity(0.0)
        self.setGraphicsEffect(self._opacity)
        self._anim = QPropertyAnimation(self._opacity, b"opacity")
        self._anim.setDuration(150)
        self._anim.setEasingCurve(QEasingCurve.Type.InOutQuad)

    def _sort_with_favourites(
        self, devices: list[AudioDevice]
    ) -> tuple[list[AudioDevice], list[AudioDevice]]:
        """Split devices into (favourites, rest), preserving order."""
        favs = [d for d in devices if self._config_mgr.is_favourite(d.id)]
        rest = [d for d in devices if not self._config_mgr.is_favourite(d.id)]
        return favs, rest

    def _add_device_rows(
        self,
        layout: QVBoxLayout,
        devices: list[AudioDevice],
        is_fav: bool = False,
    ) -> None:
        for dev in devices:
            row = DeviceRow(
                dev, self._icon_mgr, self._config_mgr,
                is_fav=is_fav, disconnected=not dev.is_connected,
            )
            row.clicked.connect(self._on_device_clicked)
            row.fav_toggled.connect(self._on_fav_toggled)
            layout.addWidget(row)
            self._rows_by_device_id[dev.id] = row

    def populate_and_show(self, anchor: QPoint, show_mode: str, widget_height: int = 40, *, reposition: bool = True) -> None:
        """Populate device list and show the dropdown near *anchor*.

        When *reposition* is False the dropdown keeps its current position
        (used for in-place refresh after a BT operation).
        """
        self._last_show_mode = show_mode
        self._last_widget_height = widget_height
        # Clear previous content
        if self.layout():
            while self.layout().count():
                item = self.layout().takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            QWidget().setLayout(self.layout())

        t = _t(self._config_mgr)
        self._rows_by_device_id.clear()
        devices = self._audio_mgr.enumerate_devices()
        active_ids = {d.id for d in devices}

        # Backfill favourite_devices metadata for any active favourite
        # that was added before the metadata feature existed.
        fav_metadata = self._config_mgr.get_favourite_devices()
        for dev in devices:
            if self._config_mgr.is_favourite(dev.id) and dev.id not in fav_metadata:
                log.info("Backfilling favourite metadata for '%s'", dev.name)
                fav_metadata[dev.id] = {
                    "name": dev.name,
                    "flow": dev.flow.value,
                    "is_bluetooth": dev.is_bluetooth,
                }
                self._config_mgr.config.favourite_devices[dev.id] = fav_metadata[dev.id]
                self._config_mgr.save()

        # Reconcile BT endpoint IDs: when a Bluetooth device reconnects it
        # often gets a new endpoint ID.  Match by extracted BT device name
        # and migrate the old favourite to the new ID.
        for dev in list(devices):
            if not dev.is_bluetooth or self._config_mgr.is_favourite(dev.id):
                continue  # skip non-BT or already-favourite devices
            # Look for a disconnected favourite with a matching BT name
            for fav_id in list(self._config_mgr.config.favourites):
                if fav_id in active_ids:
                    continue  # this favourite is still active, skip
                meta = fav_metadata.get(fav_id)
                if not meta or not meta.get("is_bluetooth"):
                    continue
                if _bt_names_match(dev.name, meta.get("name", "")):
                    log.info(
                        "Reconciling BT favourite: '%s' (old ID %s) → '%s' (new ID %s)",
                        meta.get("name"), fav_id, dev.name, dev.id,
                    )
                    self._config_mgr.migrate_favourite_id(fav_id, dev.id, dev.name)
                    # Update local tracking so ghost logic doesn't re-add
                    active_ids.add(dev.id)
                    fav_metadata = self._config_mgr.get_favourite_devices()
                    break

        # Add ghost entries for disconnected favourites
        for fav_id in list(self._config_mgr.config.favourites):
            if fav_id not in active_ids:
                meta = fav_metadata.get(fav_id)
                if meta:
                    ghost = AudioDevice(
                        id=fav_id,
                        name=meta.get("name", "Unknown"),
                        flow=DeviceFlow(meta.get("flow", "output")),
                        is_default=False,
                        is_bluetooth=meta.get("is_bluetooth", False),
                        is_connected=False,
                    )
                    devices.append(ghost)
                    log.debug("Added ghost entry for disconnected favourite: '%s'",
                              meta.get("name", fav_id))
                else:
                    log.debug("Skipping ghost for favourite %s (no metadata)", fav_id)

        outputs = [d for d in devices if d.flow == DeviceFlow.OUTPUT]
        inputs = [d for d in devices if d.flow == DeviceFlow.INPUT]

        # Pre-calculate size to determine grow direction
        row_count = 0
        if show_mode in ("output", "both"):
            row_count += len(outputs) + 1  # +1 for header
        if show_mode in ("input", "both"):
            row_count += len(inputs) + 1
        est_height = min(row_count * 38 + 50, 420)
        screen_geo = QApplication.primaryScreen().availableGeometry()
        grows_up = (anchor.y() + est_height) > screen_geo.bottom()
        # When grows_up: favourites at bottom (closer to widget) = closer to cursor
        # When grows_down: favourites at top (closer to widget) = closer to cursor
        favs_at_end = grows_up

        container = QWidget()
        container.setObjectName("Dropdown")
        container.setStyleSheet(_dropdown_stylesheet(t))
        v_layout = QVBoxLayout(container)
        v_layout.setContentsMargins(4, 8, 4, 8)
        v_layout.setSpacing(0)

        def _add_section(
            label: str, device_list: list[AudioDevice], need_sep: bool
        ) -> None:
            if not device_list:
                return
            if need_sep:
                sep = QWidget()
                sep.setFixedHeight(1)
                sep.setStyleSheet(f"background-color: {t['border']};")
                v_layout.addWidget(sep)
            header = QLabel(label)
            header.setObjectName("SectionHeader")
            header.setFont(QFont(_FONT_FAMILY, 9, QFont.Weight.Bold))
            v_layout.addWidget(header)

            favs, rest = self._sort_with_favourites(device_list)
            if favs_at_end:
                # Favourites at bottom (dropdown grows up, widget at bottom)
                self._add_device_rows(v_layout, rest)
                if favs and rest:
                    fav_sep = QWidget()
                    fav_sep.setFixedHeight(1)
                    fav_sep.setStyleSheet(
                        f"background-color: {t['border']}; margin: 2px 12px;"
                    )
                    v_layout.addWidget(fav_sep)
                if favs:
                    self._add_device_rows(v_layout, favs, is_fav=True)
            else:
                # Favourites at top (dropdown grows down, widget at top)
                if favs:
                    self._add_device_rows(v_layout, favs, is_fav=True)
                    if rest:
                        fav_sep = QWidget()
                        fav_sep.setFixedHeight(1)
                        fav_sep.setStyleSheet(
                            f"background-color: {t['border']}; margin: 2px 12px;"
                        )
                        v_layout.addWidget(fav_sep)
                self._add_device_rows(v_layout, rest)

        added = False
        if show_mode in ("output", "both") and outputs:
            _add_section("Output Devices", outputs, False)
            added = True
        if show_mode in ("input", "both") and inputs:
            _add_section("Input Devices", inputs, added)

        if not outputs and not inputs:
            empty = QLabel("No audio devices found")
            empty.setStyleSheet(f"color: {t['fg_dim']}; padding: 16px;")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            v_layout.addWidget(empty)

        # Scroll area
        scroll = QScrollArea()
        scroll.setWidget(container)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setMaximumHeight(400)
        scroll.setStyleSheet(_scroll_stylesheet(t))

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

        # Size and position
        container.adjustSize()
        width = max(container.sizeHint().width() + 16, 280)
        height = min(container.sizeHint().height() + 16, 420)
        self.setFixedSize(width, height)

        if reposition:
            x = anchor.x()
            if grows_up:
                y = anchor.y() - height - widget_height - 4
            else:
                y = anchor.y()
            if x + width > screen_geo.right():
                x = screen_geo.right() - width
            if y < screen_geo.top():
                y = screen_geo.top()
            self.move(x, y)

            # Fade in
            self._opacity.setOpacity(0.0)
            self.show()
            self._anim.setStartValue(0.0)
            self._anim.setEndValue(1.0)
            self._anim.start()

    def _on_device_clicked(self, device: AudioDevice) -> None:
        if self._bt_busy:
            return  # ignore clicks while a BT op is running

        # BT connect: disconnected BT favourite → keep dropdown open
        if not device.is_connected and device.is_bluetooth:
            self._bt_busy = True
            self._set_row_status(device.id, "Connecting\u2026")
            self.bt_connect_requested.emit(device)
            return

        # BT disconnect: default BT device → keep dropdown open
        if device.is_default and device.is_bluetooth:
            self._bt_busy = True
            self._set_row_status(device.id, "Disconnecting\u2026")
            self.bt_disconnect_requested.emit(device)
            return

        # Normal device selection — emit and close
        self.device_selected.emit(device)
        self._fade_out_and_close()

    def _set_row_status(self, device_id: str, text: str) -> None:
        """Update the name label on a device row (e.g. 'Connecting…')."""
        row = self._rows_by_device_id.get(device_id)
        if row:
            row.set_name_text(text)

    def show_bt_result(self, device_id: str, success: bool, action: str) -> None:
        """Handle BT operation result.

        On success → repopulate with fresh device data so state is current.
        On failure → show failure text on the row.
        """
        self._bt_busy = False
        if success:
            self._repopulate()
        else:
            label = "Connection failed" if action == "connect" else "Disconnect failed"
            self._set_row_status(device_id, label)

    def _repopulate(self) -> None:
        """Re-populate the dropdown in-place with fresh device data."""
        self.populate_and_show(
            self.pos(),
            self._last_show_mode,
            self._last_widget_height,
            reposition=False,
        )

    def _on_fav_toggled(self, device: AudioDevice) -> None:
        self.favourite_toggled.emit(device)

    def _fade_out_and_close(self) -> None:
        self._anim.setStartValue(1.0)
        self._anim.setEndValue(0.0)
        self._anim.finished.connect(self.close)
        self._anim.start()

    def hideEvent(self, event: object) -> None:
        self.closed.emit()
        super().hideEvent(event)


# ---------------------------------------------------------------------------
# _BodyWidget — widget body that paints a volume bar in its own paint cycle
# ---------------------------------------------------------------------------
class _BodyWidget(QWidget):
    """Widget body that can also paint a volume overlay bar.

    By painting the volume bar in the same paintEvent as the background,
    we avoid the parent-child repaint flicker of a separate overlay widget.
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._vol_level: float = 0.0
        self._vol_opacity: float = 0.0
        self._vol_accent = QColor("#89b4fa")

        self._vol_hide_timer = QTimer(self)
        self._vol_hide_timer.setSingleShot(True)
        self._vol_hide_timer.timeout.connect(self._vol_start_fade)

        self._vol_fade_timer = QTimer(self)
        self._vol_fade_timer.setInterval(30)
        self._vol_fade_timer.timeout.connect(self._vol_tick_fade)

    def show_volume(self, level: float, theme: dict[str, str]) -> None:
        self._vol_level = max(0.0, min(1.0, level))
        self._vol_accent = QColor(theme["accent"])
        self._vol_opacity = 1.0
        self._vol_fade_timer.stop()
        self._vol_hide_timer.start(1500)
        self.update()

    @property
    def volume_active(self) -> bool:
        return self._vol_opacity > 0.01

    def _vol_start_fade(self) -> None:
        self._vol_fade_timer.start()

    def _vol_tick_fade(self) -> None:
        self._vol_opacity -= 0.06
        if self._vol_opacity <= 0.0:
            self._vol_opacity = 0.0
            self._vol_fade_timer.stop()
        self.update()

    def paintEvent(self, event: QPaintEvent) -> None:
        # Custom QWidget subclasses must explicitly paint their stylesheet
        painter = QPainter(self)
        opt = QStyleOption()
        opt.initFrom(self)
        self.style().drawPrimitive(QStyle.PrimitiveElement.PE_Widget, opt, painter, self)

        if self._vol_opacity > 0.01:
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            painter.setOpacity(self._vol_opacity)
            path = QPainterPath()
            path.addRoundedRect(QRectF(self.rect()), _RADIUS, _RADIUS)
            painter.setClipPath(path)
            fill_w = int(self.width() * self._vol_level)
            bar_color = QColor(self._vol_accent)
            bar_color.setAlpha(80)
            painter.fillRect(0, 0, fill_w, self.height(), bar_color)

        painter.end()


# ---------------------------------------------------------------------------
# _BluetoothWorker — runs BT connect/disconnect off the UI thread
# ---------------------------------------------------------------------------
class _BluetoothWorker(QObject):
    """Performs Bluetooth connect/disconnect in a background thread."""

    finished = pyqtSignal(bool, str)  # success, action ("connect" or "disconnect")

    def __init__(self, device_name: str, action: str) -> None:
        super().__init__()
        self._device_name = device_name
        self._action = action  # "connect" or "disconnect"

    def run(self) -> None:
        if self._action == "connect":
            ok = bluetooth_connect(self._device_name)
        else:
            ok = bluetooth_disconnect(self._device_name)
        self.finished.emit(ok, self._action)


# ---------------------------------------------------------------------------
# AudioFlipWidget — the main always-on-top widget
# ---------------------------------------------------------------------------
class AudioFlipWidget(QWidget):
    """Compact, draggable, always-on-top audio device switcher widget."""

    def __init__(
        self,
        audio_mgr: AudioManager,
        icon_mgr: IconManager,
        config_mgr: ConfigManager,
    ) -> None:
        flags = (
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.Tool
            | Qt.WindowType.WindowStaysOnTopHint
        )
        super().__init__(None, flags)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setObjectName("WidgetBody")

        self._audio_mgr = audio_mgr
        self._icon_mgr = icon_mgr
        self._config_mgr = config_mgr
        self._drag_offset: QPoint | None = None
        self._drag_origin: QPoint | None = None
        self._dragged: bool = False
        self._dropdown: DeviceDropdown | None = None
        self._dropdown_closed_at: float = 0.0
        self._ctx_menu_closed_at: float = 0.0
        self._last_default_id: str | None = None  # track for flash detection
        self._flash_alpha: float = 0.0
        self._bt_thread: QThread | None = None
        self._bt_worker: _BluetoothWorker | None = None
        self._bt_pending_device_id: str | None = None  # old device ID (may be stale after reconnect)
        self._bt_pending_device_name: str | None = None  # device name for name-based fallback
        self._bt_active_device_id: str | None = None  # device ID the dropdown is showing status for

        # Build body container
        self._body = _BodyWidget(self)
        self._body.setObjectName("WidgetBody")

        body_layout = QHBoxLayout(self._body)
        body_layout.setContentsMargins(10, 6, 14, 6)
        body_layout.setSpacing(8)

        self._icon_label = QLabel()
        self._icon_label.setFixedSize(22, 22)
        body_layout.addWidget(self._icon_label)

        self._name_label = QLabel()
        self._name_label.setFont(QFont(_FONT_FAMILY, 10))
        self._name_label.setMaximumWidth(160)
        body_layout.addWidget(self._name_label)

        outer = QHBoxLayout(self)
        outer.setContentsMargins(3, 3, 3, 3)  # room for flash glow border
        outer.addWidget(self._body)

        # Flash animation state
        self._flash_timer = QTimer(self)
        self._flash_timer.setInterval(30)  # ~33 fps
        self._flash_timer.timeout.connect(self._tick_flash)
        self._flash_phase: float = 0.0

        # Apply theme
        self._apply_theme()

        # Restore position and ensure it's on a visible screen
        pos = self._config_mgr.config.position
        self.move(pos.get("x", 100), pos.get("y", 100))
        self._ensure_on_screen()

        # Re-check position when a monitor is disconnected
        app = QApplication.instance()
        if app is not None:
            app.screenRemoved.connect(self._on_screen_removed)

        # Apply always-on-top from config
        self._apply_always_on_top(self._config_mgr.config.always_on_top)

        # Refresh timer (fallback polling every 2s for device changes)
        self._poll_timer = QTimer(self)
        self._poll_timer.timeout.connect(self._refresh_display)
        self._poll_timer.start(2000)

        # Topmost re-assertion timer (500ms)
        self._topmost_timer = QTimer(self)
        self._topmost_timer.timeout.connect(self._reassert_topmost)
        if self._config_mgr.config.always_on_top:
            self._topmost_timer.start(500)

        # Register for event-driven updates
        self._audio_mgr.register_change_callback(self._on_device_change_com)

        # System tray icon
        self._tray = SystemTrayIcon(self._icon_mgr, self._config_mgr, self)
        self._tray.show()

        # Initial display
        self._refresh_display()

    # --- Theme -------------------------------------------------------------

    def _apply_theme(self) -> None:
        """Apply the current theme to the widget body and labels."""
        t = _t(self._config_mgr)
        self._body.setStyleSheet(_widget_stylesheet(t))
        self._name_label.setStyleSheet(
            f"color: {t['fg']}; background: transparent;"
        )
        self._icon_mgr.set_tint(QColor(t["icon_tint"]))
        self.update()

    # --- Screen / position guard -------------------------------------------

    def _ensure_on_screen(self) -> None:
        """If the widget centre is not within any visible screen, move it."""
        centre = self.geometry().center()
        for screen in QApplication.screens():
            if screen.availableGeometry().contains(centre):
                return  # visible — nothing to do
        # Off-screen: move to centre of primary screen
        primary = QApplication.primaryScreen()
        if primary is None:
            return
        geo = primary.availableGeometry()
        x = geo.x() + (geo.width() - self.width()) // 2
        y = geo.y() + (geo.height() - self.height()) // 2
        log.info("Widget off-screen — relocating to (%d, %d)", x, y)
        self.move(x, y)
        self._config_mgr.set_position(x, y)

    def _on_screen_removed(self, _screen: object) -> None:
        """Called when a monitor is disconnected."""
        # Short delay so Qt can update its screen list before we check
        QTimer.singleShot(500, self._ensure_on_screen)

    def _move_to_screen(self) -> None:
        """Move the widget to the centre of the current primary screen."""
        primary = QApplication.primaryScreen()
        if primary is None:
            return
        geo = primary.availableGeometry()
        x = geo.x() + (geo.width() - self.width()) // 2
        y = geo.y() + (geo.height() - self.height()) // 2
        self.move(x, y)
        self._config_mgr.set_position(x, y)
        self.show()
        self._apply_always_on_top(self._config_mgr.config.always_on_top)

    # --- Border flash animation -------------------------------------------

    def _start_flash(self) -> None:
        """Start a brief white border glow if enabled."""
        if not self._config_mgr.config.flash_on_change:
            return
        self._flash_phase = 0.0
        self._flash_timer.start()

    def _tick_flash(self) -> None:
        """Advance flash animation. Fades 1.0 -> 0.0 over ~500ms."""
        self._flash_phase += 0.06  # ~30 ticks * 0.06 ≈ 1.8 radians peak
        self._flash_alpha = max(0.0, 1.0 - self._flash_phase)
        if self._flash_alpha <= 0.0:
            self._flash_timer.stop()
            self._flash_alpha = 0.0
        self.update()

    def paintEvent(self, event: QPaintEvent) -> None:
        """Draw the flash border glow when active."""
        super().paintEvent(event)
        if self._flash_alpha > 0.01:
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            color = QColor(255, 255, 255, int(self._flash_alpha * 255))
            pen = QPen(color, 3.0)
            painter.setPen(pen)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            # Draw slightly outside the body so it's visible over the existing border
            r = self._body.geometry().adjusted(-1, -1, 1, 1)
            painter.drawRoundedRect(r, _RADIUS + 1, _RADIUS + 1)
            painter.end()

    # --- Display refresh ---------------------------------------------------

    def _refresh_display(self) -> None:
        """Update the widget icon and label to reflect the current default device."""
        mode = self._config_mgr.config.show_mode
        device: AudioDevice | None = None
        if mode in ("output", "both"):
            device = self._audio_mgr.get_default_output()
        if device is None and mode in ("input", "both"):
            device = self._audio_mgr.get_default_input()

        t = _t(self._config_mgr)

        if device is None:
            self._icon_label.setPixmap(
                self._icon_mgr.get_icon("audio", 20).pixmap(QSize(20, 20))
            )
            self._name_label.setText("No device")
            self._last_default_id = None
            return

        # Detect device change -> flash
        if self._last_default_id is not None and device.id != self._last_default_id:
            self._start_flash()
        self._last_default_id = device.id

        override = self._config_mgr.get_icon_override(device.id)
        icon_key = override or match_icon_for_name(device.name)
        self._icon_label.setPixmap(
            self._icon_mgr.get_icon(icon_key, 20).pixmap(QSize(20, 20))
        )

        metrics = self._name_label.fontMetrics()
        elided = metrics.elidedText(
            device.name,
            Qt.TextElideMode.ElideRight,
            self._name_label.maximumWidth(),
        )
        self._name_label.setText(elided)
        self._name_label.setToolTip("")

    def _on_device_change_com(self) -> None:
        """Called from COM thread — schedule a Qt-safe refresh."""
        QTimer.singleShot(0, self._refresh_display)

    # --- Mouse interaction -------------------------------------------------

    def _reset_drag(self) -> None:
        """Clear all drag state."""
        self._drag_offset = None
        self._drag_origin = None
        self._dragged = False

    def mousePressEvent(self, event: QMouseEvent) -> None:
        # Always clear stale drag state first (popups can swallow releases)
        self._reset_drag()
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_origin = event.globalPosition().toPoint()
            self._drag_offset = self._drag_origin - self.pos()
            self._dragged = False
        elif event.button() == Qt.MouseButton.RightButton:
            self._show_context_menu(event.globalPosition().toPoint())

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        # Check physical button state to avoid stale drag from popup interactions
        if not (QApplication.mouseButtons() & Qt.MouseButton.LeftButton):
            self._reset_drag()
            return
        if self._drag_offset is None:
            return
        if not self._dragged:
            delta = event.globalPosition().toPoint() - self._drag_origin
            if abs(delta.x()) > 4 or abs(delta.y()) > 4:
                self._dragged = True
            else:
                return  # don't move until past the drag threshold
        new_pos = event.globalPosition().toPoint() - self._drag_offset
        self.move(new_pos)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            if self._drag_offset is not None:
                if not self._dragged:
                    self._open_dropdown()
                self._config_mgr.set_position(self.x(), self.y())
            self._reset_drag()

    def _open_dropdown(self) -> None:
        """Show the device dropdown below the widget."""
        # If the dropdown just closed (same click dismissed it), don't reopen
        if time.monotonic() - self._dropdown_closed_at < 0.3:
            return

        if self._dropdown is not None:
            self._dropdown.close()
            self._dropdown.deleteLater()

        self._dropdown = DeviceDropdown(
            self._audio_mgr, self._icon_mgr, self._config_mgr
        )
        self._dropdown.device_selected.connect(self._on_device_selected)
        self._dropdown.bt_connect_requested.connect(self._bt_connect_and_switch)
        self._dropdown.bt_disconnect_requested.connect(self._bt_disconnect)
        self._dropdown.favourite_toggled.connect(self._on_fav_toggled_dropdown)
        self._dropdown.closed.connect(self._on_dropdown_closed)

        anchor = self.mapToGlobal(QPoint(0, self.height() + 4))
        self._dropdown.populate_and_show(
            anchor, self._config_mgr.config.show_mode, self.height()
        )

    def _on_device_selected(self, device: AudioDevice) -> None:
        """Handle user clicking a non-BT device in the dropdown (normal switch)."""
        success = self._audio_mgr.set_default_device(device.id)
        if success:
            self._refresh_display()
        else:
            log.warning("Failed to switch to device: %s", device.name)

    def _bt_connect_and_switch(self, device: AudioDevice) -> None:
        """Kick off BT connect in a background thread."""
        if self._bt_thread is not None:
            return  # already running a BT operation
        self._bt_pending_device_id = device.id
        self._bt_pending_device_name = device.name
        self._bt_active_device_id = device.id
        self._name_label.setText("Connecting\u2026")

        self._bt_worker = _BluetoothWorker(device.name, "connect")
        self._bt_thread = QThread()
        self._bt_worker.moveToThread(self._bt_thread)
        self._bt_thread.started.connect(self._bt_worker.run)
        self._bt_worker.finished.connect(self._on_bt_finished)
        self._bt_worker.finished.connect(self._bt_thread.quit)
        self._bt_thread.finished.connect(self._bt_cleanup)
        self._bt_thread.start()

    def _bt_disconnect(self, device: AudioDevice) -> None:
        """Kick off BT disconnect in a background thread."""
        if self._bt_thread is not None:
            return
        self._bt_pending_device_id = None
        self._bt_pending_device_name = None
        self._bt_active_device_id = device.id
        self._name_label.setText("Disconnecting\u2026")

        self._bt_worker = _BluetoothWorker(device.name, "disconnect")
        self._bt_thread = QThread()
        self._bt_worker.moveToThread(self._bt_thread)
        self._bt_thread.started.connect(self._bt_worker.run)
        self._bt_worker.finished.connect(self._on_bt_finished)
        self._bt_worker.finished.connect(self._bt_thread.quit)
        self._bt_thread.finished.connect(self._bt_cleanup)
        self._bt_thread.start()

    def _on_bt_finished(self, success: bool, action: str) -> None:
        """Handle BT worker completion (runs on main thread via signal)."""
        device_id = self._bt_active_device_id

        # Update dropdown row status
        if self._dropdown and self._dropdown.isVisible():
            if action == "connect" and success:
                self._dropdown._set_row_status(device_id or "", "Connected \u2014 switching\u2026")
            elif not success:
                # Show failure on dropdown, it will auto-close after 2s
                self._dropdown.show_bt_result(device_id or "", False, action)

        if action == "connect" and success:
            # Wait for Windows to register the new endpoint, then switch
            QTimer.singleShot(1500, self._set_pending_bt_device)
        elif action == "disconnect" and success:
            # Close dropdown and refresh after a brief pause
            if self._dropdown and self._dropdown.isVisible():
                self._dropdown.show_bt_result(device_id or "", True, action)
            QTimer.singleShot(500, self._refresh_display)
        else:
            # Failure — show "Failed" on widget label, refresh after a pause
            fail_label = "Connect failed" if action == "connect" else "Disconnect failed"
            self._name_label.setText(fail_label)
            self._bt_pending_device_id = None
            self._bt_pending_device_name = None
            self._bt_active_device_id = None
            QTimer.singleShot(2000, self._refresh_display)

    def _set_pending_bt_device(self) -> None:
        """Set the BT device as default after it has reconnected.

        The old device ID may no longer exist (BT devices often get new
        endpoint IDs on reconnect), so we fall back to name-based matching.
        """
        dev_id = self._bt_pending_device_id
        dev_name = self._bt_pending_device_name
        self._bt_pending_device_id = None
        self._bt_pending_device_name = None
        dropdown_device_id = self._bt_active_device_id
        self._bt_active_device_id = None

        if dev_id:
            # Try the stored ID first
            ok = self._audio_mgr.set_default_device(dev_id)
            if not ok and dev_name:
                # ID is stale — find the new endpoint by name
                log.info("Old device ID failed, searching by name: '%s'", dev_name)
                new_dev = self._find_bt_device_by_name(dev_name)
                if new_dev:
                    log.info("Found new endpoint: '%s' (%s)", new_dev.name, new_dev.id)
                    ok = self._audio_mgr.set_default_device(new_dev.id)
                else:
                    log.warning("No active BT endpoint found matching '%s'", dev_name)
            if ok:
                if self._dropdown and self._dropdown.isVisible():
                    self._dropdown.show_bt_result(dropdown_device_id or "", True, "connect")
            else:
                if self._dropdown and self._dropdown.isVisible():
                    self._dropdown.show_bt_result(dropdown_device_id or "", False, "connect")
                self._name_label.setText("Switch failed")
                QTimer.singleShot(2000, self._refresh_display)
                return

        self._refresh_display()

    def _find_bt_device_by_name(self, name: str) -> AudioDevice | None:
        """Find an active BT audio device whose name matches the given name."""
        for dev in self._audio_mgr.enumerate_devices():
            if dev.is_bluetooth and _bt_names_match(dev.name, name):
                return dev
        return None

    def _bt_cleanup(self) -> None:
        """Clean up the BT worker thread."""
        if self._bt_worker:
            self._bt_worker.deleteLater()
            self._bt_worker = None
        if self._bt_thread:
            self._bt_thread.deleteLater()
            self._bt_thread = None

    def _on_fav_toggled_dropdown(self, device: AudioDevice) -> None:
        """Toggle favourite for a device, then refresh the dropdown."""
        self._config_mgr.toggle_favourite(
            device.id, device.name, device.flow.value, device.is_bluetooth,
        )
        # Re-open to reflect change — reset timestamp so the re-open isn't suppressed
        self._dropdown_closed_at = 0.0
        self._open_dropdown()

    def _on_dropdown_closed(self) -> None:
        """Record when the dropdown closes so we can suppress immediate reopen."""
        self._dropdown_closed_at = time.monotonic()
        self._reset_drag()

    # --- Volume scroll -----------------------------------------------------

    def wheelEvent(self, event: QWheelEvent) -> None:
        """Scroll wheel adjusts the default device volume."""
        delta = event.angleDelta().y()
        if delta == 0:
            return
        mode = self._config_mgr.config.show_mode
        flow = DeviceFlow.INPUT if mode == "input" else DeviceFlow.OUTPUT

        # Use cached level if overlay is active to avoid API rounding jitter
        if self._body.volume_active:
            current = self._body._vol_level
        else:
            current = self._audio_mgr.get_default_volume(flow)
            if current is None:
                return

        step = (delta / 120) * 0.02  # 2% per scroll notch
        new_level = max(0.0, min(1.0, current + step))
        if self._audio_mgr.set_default_volume(new_level, flow):
            self._body.show_volume(new_level, _t(self._config_mgr))

    # --- Context menu (shared between widget and tray) ---------------------

    def build_context_menu(self) -> QMenu:
        """Build the full context menu. Used by both widget right-click and tray."""
        t = _t(self._config_mgr)
        menu = QMenu(self)
        menu.setStyleSheet(_menu_stylesheet(t))

        # Always on top
        aot_action = QAction("Always on Top", self)
        aot_action.setCheckable(True)
        aot_action.setChecked(self._config_mgr.config.always_on_top)
        aot_action.triggered.connect(self._toggle_always_on_top)
        menu.addAction(aot_action)

        # Flash on change
        flash_action = QAction("Flash on Change", self)
        flash_action.setCheckable(True)
        flash_action.setChecked(self._config_mgr.config.flash_on_change)
        flash_action.triggered.connect(self._toggle_flash)
        menu.addAction(flash_action)

        menu.addSeparator()

        # Show mode submenu
        mode_menu = menu.addMenu("Show Devices")
        mode_menu.setStyleSheet(menu.styleSheet())
        current_mode = self._config_mgr.config.show_mode
        for label, mode_val in [
            ("Output Only", "output"),
            ("Input Only", "input"),
            ("Both", "both"),
        ]:
            action = QAction(label, self)
            action.setCheckable(True)
            action.setChecked(current_mode == mode_val)
            action.triggered.connect(
                lambda checked, m=mode_val: self._set_show_mode(m)
            )
            mode_menu.addAction(action)

        # Change icon for current device
        change_icon_menu = menu.addMenu("Change Icon")
        change_icon_menu.setStyleSheet(menu.styleSheet())
        mode = self._config_mgr.config.show_mode
        device = (
            self._audio_mgr.get_default_output()
            if mode != "input"
            else self._audio_mgr.get_default_input()
        )
        if device:
            for icon_key in ICON_TYPES:
                icon_action = QAction(
                    self._icon_mgr.get_icon(icon_key, 16),
                    icon_key.replace("_", " ").title(),
                    self,
                )
                icon_action.triggered.connect(
                    lambda checked, d=device.id, k=icon_key: self._change_device_icon(
                        d, k
                    )
                )
                change_icon_menu.addAction(icon_action)

        menu.addSeparator()

        # Theme submenu
        theme_menu = menu.addMenu("Theme")
        theme_menu.setStyleSheet(menu.styleSheet())
        current_theme = self._config_mgr.config.theme
        for theme_name in VALID_THEMES:
            t_action = QAction(theme_name.capitalize(), self)
            t_action.setCheckable(True)
            t_action.setChecked(current_theme == theme_name)
            t_action.triggered.connect(
                lambda checked, tn=theme_name: self._set_theme(tn)
            )
            theme_menu.addAction(t_action)

        menu.addSeparator()

        # Start with Windows
        startup_action = QAction("Start with Windows", self)
        startup_action.setCheckable(True)
        startup_action.setChecked(self._config_mgr.config.start_with_windows)
        startup_action.triggered.connect(self._toggle_startup)
        menu.addAction(startup_action)

        # Move to Screen
        move_action = QAction("Move to Screen", self)
        move_action.triggered.connect(self._move_to_screen)
        menu.addAction(move_action)

        menu.addSeparator()

        # Quit
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(QApplication.quit)
        menu.addAction(quit_action)

        return menu

    def _show_context_menu(self, pos: QPoint) -> None:
        # If the context menu just closed (same click dismissed it), don't reopen
        if time.monotonic() - self._ctx_menu_closed_at < 0.3:
            return

        # Clean up previous menu if any
        if hasattr(self, "_ctx_menu") and self._ctx_menu is not None:
            self._ctx_menu.deleteLater()

        # Pause topmost re-assertion while menu is open (prevents the widget
        # from being pushed back above the menu/submenus every 500ms)
        self._topmost_timer.stop()

        self._ctx_menu = self.build_context_menu()
        self._ctx_menu.aboutToHide.connect(self._restore_topmost_after_menu)

        # Install event filter so menus and submenus get elevated to TOPMOST
        self._ctx_menu.installEventFilter(self)
        for action in self._ctx_menu.actions():
            submenu = action.menu()
            if submenu:
                submenu.installEventFilter(self)

        # Position above or below the widget, like the dropdown
        screen_geo = QApplication.primaryScreen().availableGeometry()
        menu_size = self._ctx_menu.sizeHint()
        anchor_below = self.mapToGlobal(QPoint(0, self.height() + 4))
        anchor_above = self.mapToGlobal(QPoint(0, -menu_size.height() - 4))
        if anchor_below.y() + menu_size.height() <= screen_geo.bottom():
            show_pos = anchor_below
        else:
            show_pos = anchor_above
        if show_pos.x() + menu_size.width() > screen_geo.right():
            show_pos.setX(screen_geo.right() - menu_size.width())
        self._ctx_menu.popup(show_pos)

    def eventFilter(self, obj: object, event: QEvent) -> bool:
        """Elevate QMenu popups to HWND_TOPMOST so they appear above this widget."""
        if isinstance(obj, QMenu) and event.type() == QEvent.Type.Show:
            QTimer.singleShot(0, lambda m=obj: self._elevate_menu(m))
        return super().eventFilter(obj, event)

    def _elevate_menu(self, menu: QMenu) -> None:
        """Make a menu window HWND_TOPMOST via Win32."""
        try:
            menu_hwnd = int(menu.winId())
            if menu_hwnd:
                _user32.SetWindowPos(
                    menu_hwnd, _HWND_TOPMOST, 0, 0, 0, 0,
                    _SWP_NOMOVE | _SWP_NOSIZE | _SWP_NOACTIVATE | _SWP_SHOWWINDOW,
                )
        except Exception:
            pass

    def _restore_topmost_after_menu(self) -> None:
        """Restart the topmost re-assertion timer after the menu closes."""
        self._ctx_menu_closed_at = time.monotonic()
        self._reset_drag()
        if self._config_mgr.config.always_on_top:
            self._topmost_timer.start(500)

    def _toggle_always_on_top(self, checked: bool) -> None:
        self._config_mgr.set_always_on_top(checked)
        self._apply_always_on_top(checked)

    def _apply_always_on_top(self, on_top: bool) -> None:
        """Apply always-on-top using Win32 SetWindowPos for true taskbar-level topmost."""
        hwnd = int(self.winId())
        if on_top:
            _user32.SetWindowPos(
                hwnd,
                _HWND_TOPMOST,
                0,
                0,
                0,
                0,
                _SWP_NOMOVE | _SWP_NOSIZE | _SWP_NOACTIVATE | _SWP_SHOWWINDOW,
            )
            if hasattr(self, "_topmost_timer"):
                self._topmost_timer.start(500)
        else:
            _user32.SetWindowPos(
                hwnd,
                _HWND_NOTOPMOST,
                0,
                0,
                0,
                0,
                _SWP_NOMOVE | _SWP_NOSIZE | _SWP_NOACTIVATE,
            )
            if hasattr(self, "_topmost_timer"):
                self._topmost_timer.stop()

    def _reassert_topmost(self) -> None:
        """Re-assert HWND_TOPMOST silently — no activate, no flash."""
        if not self._config_mgr.config.always_on_top:
            return
        hwnd = int(self.winId())
        _user32.SetWindowPos(
            hwnd,
            _HWND_TOPMOST,
            0,
            0,
            0,
            0,
            _SWP_NOMOVE | _SWP_NOSIZE | _SWP_NOACTIVATE,
        )

    def _set_show_mode(self, mode: str) -> None:
        self._config_mgr.set_show_mode(mode)
        self._refresh_display()

    def _change_device_icon(self, device_id: str, icon_key: str) -> None:
        self._config_mgr.set_icon_override(device_id, icon_key)
        self._apply_theme()
        self._refresh_display()

    def _set_theme(self, theme: str) -> None:
        self._config_mgr.set_theme(theme)
        self._apply_theme()
        self._refresh_display()

    def _toggle_flash(self, checked: bool) -> None:
        self._config_mgr.set_flash_on_change(checked)

    def _toggle_startup(self, checked: bool) -> None:
        self._config_mgr.set_start_with_windows(checked)

    # --- Cleanup -----------------------------------------------------------

    def closeEvent(self, event: object) -> None:
        self._audio_mgr.unregister_change_callback()
        self._poll_timer.stop()
        self._topmost_timer.stop()
        self._flash_timer.stop()
        if self._bt_thread and self._bt_thread.isRunning():
            self._bt_thread.quit()
            self._bt_thread.wait(3000)
        if self._tray:
            self._tray.hide()
        if self._dropdown:
            self._dropdown.close()
        super().closeEvent(event)


# ---------------------------------------------------------------------------
# SystemTrayIcon
# ---------------------------------------------------------------------------
class SystemTrayIcon(QSystemTrayIcon):
    """System tray icon with the same context menu as the widget."""

    def __init__(
        self,
        icon_mgr: IconManager,
        config_mgr: ConfigManager,
        widget: AudioFlipWidget,
    ) -> None:
        tray_icon = icon_mgr.get_icon("audio", 64)
        super().__init__(tray_icon, widget)
        self._widget = widget
        self._config_mgr = config_mgr
        self.setToolTip("audioflip")
        self.activated.connect(self._on_activated)

    def _on_activated(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            # Left-click: toggle widget visibility
            if self._widget.isVisible():
                self._widget.hide()
            else:
                self._widget.show()
                self._widget._apply_always_on_top(
                    self._config_mgr.config.always_on_top
                )
        elif reason == QSystemTrayIcon.ActivationReason.Context:
            menu = self._widget.build_context_menu()
            menu.exec(QCursor.pos())
