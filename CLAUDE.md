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
.venv\Scripts\python.exe tests\test_name_matching.py
```

There is no linter or formatter configured, and test coverage is limited to pure functions - fullscreen geometry, Bluetooth endpoint name matching, Bluetooth filter-path parsing, icon keyword matching. Anything touching Win32, COM or Qt is verified manually: run the app and exercise the affected path.

When a native crash happens the process dies silently, because the packaged build has no console. `faulthandler` writes the stack of every thread to `%APPDATA%\audioflip\crash.log` - check there first when the app vanishes and `audioflip.log` simply stops mid-sentence.

Runtime state lives outside the repo:

- `%APPDATA%\audioflip\config.json` - all settings
- `%APPDATA%\audioflip\audioflip.log` - rotating log, 1 MB x 4. The packaged exe is built with `console=False`, so this file is the only way to see output from a built binary

## Architecture

Windows-only PyQt6 desktop widget for switching the default audio device and driving Bluetooth audio. Six modules, all under `src/`.

`main.py` wires three managers together and hands them to `AudioFlipWidget`, which owns essentially all behaviour. `QApplication.setQuitOnLastWindowClosed(False)` is deliberate: hiding the widget from the tray must not exit the process.

- `config.py` - `ConfigManager` singleton over `config.json`. Every setter saves immediately; there is no dirty/flush model
- `audio_manager.py` - Core Audio through `pycaw`/`comtypes`. Enumeration, default-device switching, volume, change callbacks
- `bluetooth.py` - raw `ctypes`. KS one-shot connect/disconnect (`ks_oneshot()`), plus `BluetoothApis.dll` discovery, SSP pairing, service enable/disable, and a PowerShell PnP fallback. Module-level state (radio handle, paired-name cache, loaded DLL) is process-global
- `icons.py` - SVG to tinted `QIcon`, keyword matching of device name to icon key
- `ui.py` - widget, dropdown, device rows, context menu, BT scan dialog, tray icon, worker threads

### Device switching

Uses the undocumented `IPolicyConfig` COM interface (CLSID `870af99c-171d-4f9e-af0d-e63df40c2bc9`), calling `SetDefaultEndpoint` for all three roles (Console, Multimedia, Communications). There is deliberately no fallback: the previous one shelled out to `Set-AudioDevice` from the third-party AudioDeviceCmdlets module, which is not a dependency and is not normally installed, so it spawned a process only to fail.

`get_default_device()` asks Windows for the default endpoint directly rather than enumerating everything and filtering - the widget polls it every 2s, so the full enumeration was several times the necessary COM work. `_read_device_props()` is shared with `enumerate_devices()` so the Bluetooth detection rules stay in one place.

**Never `ctypes.cast()` a comtypes pointer - use `QueryInterface`.** `cast()` creates a second comtypes pointer over the single reference you were given, and both call `Release()` when collected. That was the cause of the long-unexplained crash: `get_default_volume()` cast an `IAudioEndpointVolume` on every 2s poll, freeing it while a pointer was still live, and `crash.log` recorded the access violation in `Release()` during garbage collection, every few days. Measured: after the cast pattern the refcount read `0xC0000001` (freed memory); with `QueryInterface` it stays exact over 500 cycles.

Refresh is deliberately doubled up: an `IMMNotificationClient` COM callback plus a 2-second poll. The COM callback fires on a COM thread, so it must marshal to Qt via `QTimer.singleShot(0, ...)` - never touch widgets directly from it.

### The three problems most of the complexity exists to solve

**1. Bluetooth detection is unreliable.** The correct signal is `PKEY_Device_EnumeratorName` (pid 24) starting with `BTH`. Intel/Dell audio controllers proxy BT audio through their own driver and report `INTELAUDIO` instead, so `enumerate_devices()` falls back to matching the endpoint name against `get_paired_device_names()`. That list is cached for 30 seconds because enumeration is polled roughly once per second.

**2. BT endpoint IDs used to change on reconnect.** The cause was the old connect mechanism: toggling A2DP/HFP services uninstalls and reinstalls the audio drivers, and the reinstalled endpoint gets a new ID, which would orphan its favourite and icon override. The KS path (problem 3) leaves the drivers alone, so IDs now stay stable, but favourites saved under a retired ID still need migrating. `_bt_name_core()` extracts the device name from inside the last parentheses (`"Headphones (Buds Pro 2)"` gives `buds pro 2`); `DeviceDropdown.populate_and_show()` matches active BT devices against disconnected favourites on that core name and calls `migrate_favourite_id()`, which moves the ID across `favourites`, `favourite_devices`, and `icon_overrides` together. Any change to favourites or icon storage must keep those three in sync.

**3. BT connect has to actually request a connection.** `BluetoothSetServiceState(enable)` on a paired-but-disconnected headset returns `ERROR_INVALID_PARAMETER` (87) because the services are already enabled, and requests nothing - so for months, connecting a device already set up in Windows almost never worked. The primary path is now `ks_oneshot()`: `AudioManager.bt_audio_filters()` walks endpoint -> `IDeviceTopology` -> connector 0 -> the Bluetooth audio kernel-streaming filter's device path, and the worker opens it and sends `IOCTL_KS_PROPERTY` with `KSPROPSETID_BtAudio` / `KSPROPERTY_ONESHOT_RECONNECT` (or `_DISCONNECT`), the request behind the Settings app's Connect button. It works unelevated. `IPart::Activate(IKsControl)` returns `E_NOINTERFACE` on Windows 11 26300, which is why it is a raw IOCTL rather than `IKsControl`. Service toggling remains as a fallback for a device whose services were disabled (no filter exists then).

The filter path is resolved on the main thread and only the string crosses to the `QThread` worker, so no COM pointer ever crosses threads. An accepted request does not mean connected: a headset that is off or out of range just never appears. So `_poll_bt_connect()` checks every 500 ms for the endpoint (stored ID first, then BT-name match) up to a 12 s deadline, switches the moment it appears, and shows 'Connection failed' at the deadline. This replaced fixed 1.5 s and 4 s guesses, which waited needlessly on fast stacks and reported nothing when a connect failed. Disconnect is the mirror image and calls `_switch_to_fallback()` *before* starting the BT work so audio moves to a non-BT device instantly.

### UI conventions worth knowing before editing `ui.py`

- Always-on-top is not the Qt flag alone. `SetWindowPos(HWND_TOPMOST)` is re-asserted every 500 ms so the widget sits above the taskbar. That timer is stopped while the context menu is open, otherwise the widget pushes back through its own menu, and restarted in `_restore_topmost_after_menu`
- The same 500 ms tick drives the fullscreen yield. `covers_monitor()` compares the foreground window against the monitor's `rcMonitor`, never `rcWork` - a maximised window covers `rcWork` and must not count as fullscreen. The widget demotes immediately but needs two consecutive clear ticks before restoring, so alt-tab does not flicker the z-order. Yielding goes through `_yield_below()`, which inserts the widget after the fullscreen window's own HWND, and never through `_apply_always_on_top(False)`, which would stop the timer and strand it demoted. `HWND_NOTOPMOST` alone is not a yield: Windows defines it as *above all non-topmost windows*, so it only leaves the topmost band and still floats over borderless-windowed fullscreen apps - most games, and Stremio. Those are ordinary non-topmost windows that are already active, so no click re-raises them above the widget. The placement is re-applied on every tick, because any activation reorders it
- Native popup windows bleed square corners through a stylesheet `border-radius`, so every popup applies a manual `QBitmap` mask via `_rounded_mask()` after any resize
- Custom `QWidget` subclasses do not paint their own stylesheet. `_BodyWidget.paintEvent` must call `style().drawPrimitive(PE_Widget, ...)` explicitly before drawing anything else
- The volume bar is painted inside `_BodyWidget.paintEvent` rather than as a child overlay, to avoid parent/child repaint flicker
- Wheel events are forwarded to `AudioFlipWidget.wheelEvent` from the dropdown viewport, the scan dialog, and every `QMenu`, so scroll-to-volume keeps working with any popup open
- `DeviceRow.mousePressEvent` calls `event.accept()` to stop clicks falling through the popup to windows behind it
- Row click semantics: disconnected BT device connects, connected BT device *that is currently default* disconnects, anything else is a plain switch
- Dropdown grow direction flips favourite ordering. When it grows upward, favourites render at the bottom so they stay nearest the cursor
- Disconnected favourites are synthesised as `AudioDevice(is_connected=False)` ghost rows from cached config metadata, drawn at 40% icon opacity. Headsets paired in Windows but disconnected get ghost rows too, from `bt_audio_filters()` (A2DP filters only, deduplicated by BT name), so they can be connected without being favourited first
- Themes are plain dicts of seven colour keys in `_THEMES`. Adding one means updating both `_THEMES` and `VALID_THEMES` in `config.py`

### The Bluetooth PowerShell fallback

When both the KS request and `BluetoothSetServiceState` fail, `bluetooth.py` falls back to `Enable-PnpDevice`/`Disable-PnpDevice`. Three things to know:

- it must be invoked with an argument list, never `shell=True`. The original built one command string routed through `cmd.exe`, which does not process backslash escapes, so `-Confirm:$false` arrived mangled and PowerShell rejected it as a string where a switch was required. The fallback therefore failed on argument binding every time, for months, while looking like it was trying
- `Enable-PnpDevice` requires elevation, so this path still fails when audioflip runs unelevated - it now says so explicitly rather than failing obscurely
- enabling is only attempted when the device is actually disabled (`ConfigManagerErrorCode` 22). `Enable-PnpDevice` on an enabled device exits 0 having done nothing, and that no-op was logged as 'BT connect via PowerShell succeeded' on every failed connect

### Position and screen handling

Two rules here are load-bearing, and both were previously bugs:

- **Visibility is tested against `screen.geometry()`, never `availableGeometry()`.** The widget is designed to float over the taskbar, so a position inside the taskbar strip is deliberate. Testing against `availableGeometry()` (which excludes the taskbar) evicted the widget on every startup for anyone who parked it there
- **Only deliberate user actions persist a position.** `_save_desired_position()` is the single caller of `ConfigManager.set_position()`, reached from drag-release and 'Move to screen'. The off-screen rescue moves the window but must never save, or the desired position is destroyed the first time its monitor is absent - which is every boot, since audioflip starts before Windows finishes enumerating displays

The authoritative restore runs from `showEvent`, not `__init__`, because the widget's size is still a placeholder during construction and any geometry check there is computed against the wrong rectangle. Display changes are debounced through one settle timer, since docking emits a burst of events.

### Single instance and build identification

It is single-instance: a launch first connects to a per-user `QLocalServer` pipe, and if a running copy answers, that copy surfaces its widget and the new process exits. The startup log line names the binary and its build time (`audioflip starting (C:\...\audioflip.exe, built ...)`). Both exist because stray old builds were once launched from Start menu search and looked like the app had lost features - the log is the first place to check which copy ran. Old builds predate the pipe, so they are not caught by the guard.

## Packaging notes

`audioflip.spec` needs `src.bluetooth` in `hiddenimports` because it is only reached through deferred imports inside functions. New modules imported the same way must be added there. `resources/*.svg` ships as `datas`, and `icons._resources_dir()` switches to `sys._MEIPASS` when frozen.
