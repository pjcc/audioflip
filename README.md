# AudioFlip

A compact, always-on-top Windows widget for switching audio devices and managing Bluetooth audio.

## Features

**Left-click** opens a device dropdown:
- Switch the system default audio device with one click
- Favourite devices pinned to the top (right-click a device row to toggle)
- Disconnected BT favourites stay visible as greyed-out ghost entries — click to connect
- Click the active BT device to disconnect

**Right-click** opens a context menu:
- Always on top — toggle `HWND_TOPMOST`
- Flash on change — border flash animation on device switch
- Show volume bar — persistent volume level bar on the widget
- Show devices — output only / input only / both
- Change icon — pick from 14 icon types for the current device
- Theme — 10 colour themes: `dark` `light` `midnight` `ocean` `forest` `sunset` `berry` `slate` `copper` `arctic`
- Start with windows — auto-launch on login
- Move to screen — reposition widget to the active monitor
- Scan for bluetooth — discover and pair new BT audio devices (opens a scan dialog, filters to audio devices, one-click SSP pairing)

**Scroll wheel** adjusts volume — works on the widget, and while the dropdown, context menu, or scan dialog are open.

**Other:**
- Draggable widget, position persisted between sessions
- System tray icon with the same context menu
- BT endpoint reconciliation — auto-migrates favourites/icons when Windows reassigns BT endpoint IDs on reconnect
- Per-device icons auto-assigned by name, customisable via right-click menu

## Requirements

- Windows 10 or 11
- Python 3.12+

## Setup

```powershell
git clone https://github.com/pjcc/audioflip.git
cd audioflip
python -m venv .venv
.venv\Scripts\python.exe -m ensurepip
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Running

```powershell
.venv\Scripts\python.exe run.py
```

## Building a Standalone .exe

```powershell
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
.venv\Scripts\python.exe -m PyInstaller audioflip.spec
```

Output: `dist/audioflip.exe` — single portable executable.

## Configuration

Stored in `%APPDATA%\audioflip\config.json`:

| Key | Default | Description |
|-----|---------|-------------|
| `always_on_top` | `true` | Keep widget above all windows |
| `position` | `{x:100, y:100}` | Widget screen position |
| `show_mode` | `"output"` | `"output"` / `"input"` / `"both"` |
| `start_with_windows` | `false` | Auto-launch on login |
| `theme` | `"dark"` | Colour theme |
| `flash_on_change` | `true` | Flash border on device switch |
| `show_volume_bar` | `false` | Persistent volume level bar |
| `icon_overrides` | `{}` | Per-device icon overrides (`device-id` → icon key) |
| `favourites` | `[]` | Pinned device IDs |
| `favourite_devices` | `{}` | Cached metadata for BT favourites |

### Icon Keys

`headphones` `earbuds` `speaker` `bluetooth` `monitor` `monitor2` `tv` `usb` `microphone` `camera` `soundbar` `mixer` `phone` `audio`

## How It Works

- **Device enumeration** — `pycaw` / `comtypes` via `IMMDeviceEnumerator`
- **Device switching** — undocumented `IPolicyConfig` COM interface (`SetDefaultEndpoint`)
- **Change notifications** — `IMMNotificationClient` callback + 2s polling fallback
- **Always-on-top** — `SetWindowPos(HWND_TOPMOST)` re-asserted every 500ms
- **Bluetooth** — `BluetoothSetServiceState` via `ctypes` for A2DP/HFP toggling; `BluetoothAuthenticateDeviceEx` for new-device pairing via SSP (background `QThread`)
- **BT reconciliation** — matches BT device name inside endpoint friendly name, auto-migrates favourites/icons to new endpoint ID
- **Logging** — `RotatingFileHandler` to `%APPDATA%\audioflip\audioflip.log`
