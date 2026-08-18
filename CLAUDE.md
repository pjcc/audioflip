# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```powershell
# Setup
python -m venv .venv
.venv\Scripts\python.exe -m ensurepip
.venv\Scripts\python.exe -m pip install -r requirements.txt

# Run from source
.venv\Scripts\python.exe run.py

# Build the single-file exe (clean build - stale build/ causes icon and datas staleness)
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
.venv\Scripts\python.exe -m PyInstaller audioflip.spec
```

```powershell
# Tests (dependency-free, no pytest required)
.venv\Scripts\python.exe tests\test_fullscreen.py
```

There is no linter or formatter configured, and test coverage is limited to pure geometry helpers. Anything touching Win32, COM or Qt is verified manually: run the app and exercise the affected path.

Runtime state lives outside the repo:

- `%APPDATA%\audioflip\config.json` - all settings
- `%APPDATA%\audioflip\audioflip.log` - rotating log, 1 MB x 4. The packaged exe is built with `console=False`, so this file is the only way to see output from a built binary

## Architecture

Windows-only PyQt6 desktop widget for switching the default audio device and driving Bluetooth audio. Six modules, all under `src/`.

`main.py` wires three managers together and hands them to `AudioFlipWidget`, which owns essentially all behaviour. `QApplication.setQuitOnLastWindowClosed(False)` is deliberate: hiding the widget from the tray must not exit the process.

- `config.py` - `ConfigManager` singleton over `config.json`. Every setter saves immediately; there is no dirty/flush model
- `audio_manager.py` - Core Audio through `pycaw`/`comtypes`. Enumeration, default-device switching, volume, change callbacks
- `bluetooth.py` - raw `ctypes` over `BluetoothApis.dll`. Discovery, SSP pairing, service enable/disable, plus a PowerShell PnP fallback. Module-level state (radio handle, paired-name cache, loaded DLL) is process-global
- `icons.py` - SVG to tinted `QIcon`, keyword matching of device name to icon key
- `ui.py` - widget, dropdown, device rows, context menu, BT scan dialog, tray icon, worker threads

### Device switching

Uses the undocumented `IPolicyConfig` COM interface (CLSID `870af99c-171d-4f9e-af0d-e63df40c2bc9`), calling `SetDefaultEndpoint` for all three roles (Console, Multimedia, Communications). The `_fallback_set_default` PowerShell path calls `Set-AudioDevice`, which requires the third-party `AudioDeviceCmdlets` module - it is not a dependency, so treat that fallback as effectively unavailable.

Refresh is deliberately doubled up: an `IMMNotificationClient` COM callback plus a 2-second poll. The COM callback fires on a COM thread, so it must marshal to Qt via `QTimer.singleShot(0, ...)` - never touch widgets directly from it.

### The three problems most of the complexity exists to solve

**1. Bluetooth detection is unreliable.** The correct signal is `PKEY_Device_EnumeratorName` (pid 24) starting with `BTH`. Intel/Dell audio controllers proxy BT audio through their own driver and report `INTELAUDIO` instead, so `enumerate_devices()` falls back to matching the endpoint name against `get_paired_device_names()`. That list is cached for 30 seconds because enumeration is polled roughly once per second.

**2. BT endpoint IDs change on reconnect.** A reconnecting BT device gets a new endpoint ID, which would orphan its favourite and icon override. `_bt_name_core()` extracts the device name from inside the last parentheses (`"Headphones (Buds Pro 2)"` gives `buds pro 2`); `DeviceDropdown.populate_and_show()` matches active BT devices against disconnected favourites on that core name and calls `migrate_favourite_id()`, which moves the ID across `favourites`, `favourite_devices`, and `icon_overrides` together. Any change to favourites or icon storage must keep those three in sync.

**3. BT connect is slow and stack-dependent.** Connect runs on a `QThread`, then a staged recovery follows: at 1.5 s try the stored ID, then BT-name match, then any-device name match, then check whether Windows auto-switched on its own; at 4 s a second retry catches slow stacks. Disconnect is the mirror image and calls `_switch_to_fallback()` *before* starting the BT work so audio moves to a non-BT device instantly. The timings are empirical, tuned against a Dell/Intel machine - do not tighten them casually.

### UI conventions worth knowing before editing `ui.py`

- Always-on-top is not the Qt flag alone. `SetWindowPos(HWND_TOPMOST)` is re-asserted every 500 ms so the widget sits above the taskbar. That timer is stopped while the context menu is open, otherwise the widget pushes back through its own menu, and restarted in `_restore_topmost_after_menu`
- The same 500 ms tick drives the fullscreen yield. `covers_monitor()` compares the foreground window against the monitor's `rcMonitor`, never `rcWork` - a maximised window covers `rcWork` and must not count as fullscreen. The widget demotes immediately but needs two consecutive clear ticks before restoring, so alt-tab does not flicker the z-order. Yielding must use `_set_topmost(False)` and not `_apply_always_on_top(False)`, which would stop the timer and strand it demoted
- Native popup windows bleed square corners through a stylesheet `border-radius`, so every popup applies a manual `QBitmap` mask via `_rounded_mask()` after any resize
- Custom `QWidget` subclasses do not paint their own stylesheet. `_BodyWidget.paintEvent` must call `style().drawPrimitive(PE_Widget, ...)` explicitly before drawing anything else
- The volume bar is painted inside `_BodyWidget.paintEvent` rather than as a child overlay, to avoid parent/child repaint flicker
- Wheel events are forwarded to `AudioFlipWidget.wheelEvent` from the dropdown viewport, the scan dialog, and every `QMenu`, so scroll-to-volume keeps working with any popup open
- `DeviceRow.mousePressEvent` calls `event.accept()` to stop clicks falling through the popup to windows behind it
- Row click semantics: disconnected BT device connects, connected BT device *that is currently default* disconnects, anything else is a plain switch
- Dropdown grow direction flips favourite ordering. When it grows upward, favourites render at the bottom so they stay nearest the cursor
- Disconnected favourites are synthesised as `AudioDevice(is_connected=False)` ghost rows from cached config metadata, drawn at 40% icon opacity
- Themes are plain dicts of seven colour keys in `_THEMES`. Adding one means updating both `_THEMES` and `VALID_THEMES` in `config.py`

### Position and screen handling

Two rules here are load-bearing, and both were previously bugs:

- **Visibility is tested against `screen.geometry()`, never `availableGeometry()`.** The widget is designed to float over the taskbar, so a position inside the taskbar strip is deliberate. Testing against `availableGeometry()` (which excludes the taskbar) evicted the widget on every startup for anyone who parked it there
- **Only deliberate user actions persist a position.** `_save_desired_position()` is the single caller of `ConfigManager.set_position()`, reached from drag-release and 'Move to screen'. The off-screen rescue moves the window but must never save, or the desired position is destroyed the first time its monitor is absent - which is every boot, since audioflip starts before Windows finishes enumerating displays

The authoritative restore runs from `showEvent`, not `__init__`, because the widget's size is still a placeholder during construction and any geometry check there is computed against the wrong rectangle. Display changes are debounced through one settle timer, since docking emits a burst of events.

## Packaging notes

`audioflip.spec` needs `src.bluetooth` in `hiddenimports` because it is only reached through deferred imports inside functions. New modules imported the same way must be added there. `resources/*.svg` ships as `datas`, and `icons._resources_dir()` switches to `sys._MEIPASS` when frozen.
