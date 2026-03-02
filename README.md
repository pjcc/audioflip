# AudioFlip

A compact, always-on-top Windows widget for switching audio devices and managing Bluetooth audio.

## Features

- **One-click device switching** — click the widget to open a dropdown of all audio endpoints, click to switch
- **Bluetooth connect/disconnect** — click a greyed-out BT favourite to connect; click the active BT device to disconnect
- **Scan for Bluetooth** — discover and pair new BT audio devices from the right-click menu
- **Volume control** — scroll wheel adjusts volume (works even while menus are open), optional persistent volume bar
- **Favourite devices** — pin devices to the top of the dropdown; disconnected BT favourites stay visible as ghost entries
- **BT endpoint reconciliation** — automatically migrates favourites when Windows reassigns BT endpoint IDs on reconnect
- **Per-device icons** — auto-assigned by name, customisable via right-click menu
- **10 colour themes** — `dark` `light` `midnight` `ocean` `forest` `sunset` `berry` `slate` `copper` `arctic`
- **Always-on-top** with Win32 `HWND_TOPMOST`, draggable, position persisted
- **System tray icon** with full context menu
- **Start with Windows** option
- **Border flash** on device change
- **Move to screen** — reposition widget to the active monitor

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
- **Bluetooth** — `BluetoothSetServiceState` via `ctypes` for A2DP/HFP toggling; `BluetoothAuthenticateDeviceEx` for new-device pairing (background `QThread`)
- **BT reconciliation** — matches BT device name inside endpoint friendly name, auto-migrates favourites/icons to new endpoint ID
- **Logging** — `RotatingFileHandler` to `%APPDATA%\audioflip\audioflip.log`
