# audioflip

A persistent, always-on-top Windows desktop widget for switching audio devices.

## Features

- Compact widget pinned to your screen showing the current audio device
- Click to open a dropdown of all input/output devices
- Switch the system default device with one click
- Auto-assigned icons based on device name (headphones, speakers, Bluetooth, etc.)
- Customisable per-device icons via right-click menu
- Favourite devices (right-click a device in the dropdown)
- Bluetooth-aware favourites — disconnected BT devices stay visible (greyed out) in the dropdown
- Click a greyed BT favourite to connect and switch; click the active BT device to disconnect
- Automatic Bluetooth endpoint ID reconciliation (handles Windows reassigning new IDs on reconnect)
- Scroll wheel to adjust volume (with visual overlay bar)
- Border flash animation on device change
- 10 colour themes
- Draggable positioning (persisted between sessions)
- Always-on-top toggle
- Start-with-Windows option
- System tray icon
- Event-driven device change detection

## Requirements

- Windows 10 or 11
- Python 3.12+

## Setup

> **Note:** Python 3.12+ does not guarantee pip is included in a new venv. The steps below use explicit venv paths throughout to avoid packages silently installing to the wrong location.
```powershell
# Clone the repo
git clone https://github.com/pjcc/audioflip.git
cd audioflip

# Create the virtual environment
python -m venv .venv

# Bootstrap pip into the venv
.venv\Scripts\python.exe -m ensurepip

# Install all dependencies into the venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Running
```powershell
.venv\Scripts\python.exe run.py
```

## Building a Standalone .exe
```powershell
.venv\Scripts\python.exe -m PyInstaller audioflip.spec
```

The output will be at `dist/audioflip.exe` — a single portable executable.

> If you get `PermissionError: Access is denied` on rebuild, the previous `audioflip.exe` is still running. Stop it first:
> ```powershell
> Stop-Process -Name "audioflip" -Force -ErrorAction SilentlyContinue
> ```

> **Why not just activate the venv and use `pip` normally?**  
> On Python 3.12+, `pip` after activation can still resolve to a user-level install rather than the venv. Using `.venv\Scripts\python.exe -m pip` is unambiguous regardless of system configuration.

## Configuration

Settings are stored in `%APPDATA%\audioflip\config.json`:

```json
{
  "always_on_top": true,
  "position": {"x": 100, "y": 100},
  "show_mode": "output",
  "start_with_windows": false,
  "icon_overrides": {},
  "theme": "dark",
  "favourites": [],
  "favourite_devices": {},
  "flash_on_change": true
}
```

| Key | Values | Description |
|-----|--------|-------------|
| `always_on_top` | `true` / `false` | Keep widget above all windows |
| `position` | `{"x": int, "y": int}` | Widget screen position |
| `show_mode` | `"output"` / `"input"` / `"both"` | Which devices to show |
| `start_with_windows` | `true` / `false` | Auto-launch on login |
| `icon_overrides` | `{"device-id": "icon-key"}` | Per-device icon overrides |
| `theme` | Theme name | UI theme |
| `favourites` | `["device-id", ...]` | Pinned devices |
| `favourite_devices` | `{"device-id": {...}}` | Cached metadata for favourites (name, flow, is_bluetooth) |
| `flash_on_change` | `true` / `false` | Flash border on device switch |

### Available Themes

`dark`, `light`, `midnight`, `ocean`, `forest`, `sunset`, `berry`, `slate`, `copper`, `arctic`

### Available Icon Keys

`headphones`, `earbuds`, `speaker`, `bluetooth`, `monitor`, `monitor2`, `tv`, `usb`, `microphone`, `camera`, `soundbar`, `mixer`, `phone`, `audio` (fallback)

## Project Structure

```
audioflip/
├── run.py                  # Dev entry point
├── requirements.txt
├── audioflip.spec          # PyInstaller build spec
├── README.md
├── resources/              # SVG icons
│   ├── headphones.svg
│   ├── earbuds.svg
│   ├── speaker.svg
│   ├── bluetooth.svg
│   ├── monitor.svg
│   ├── monitor2.svg
│   ├── tv.svg
│   ├── usb.svg
│   ├── microphone.svg
│   ├── camera.svg
│   ├── soundbar.svg
│   ├── mixer.svg
│   ├── phone.svg
│   ├── audio.svg           # Fallback icon
│   └── checkmark.svg
└── src/
    ├── __init__.py
    ├── main.py             # Application entry point
    ├── config.py           # Configuration management
    ├── audio_manager.py    # Windows Core Audio integration
    ├── bluetooth.py        # Bluetooth connect/disconnect via Windows API
    ├── icons.py            # Icon matching and rendering
    └── ui.py               # PyQt6 widget, dropdown, context menu
```

## How It Works

- **Device enumeration**: Uses `pycaw` and `comtypes` to query `IMMDeviceEnumerator` for active audio endpoints.
- **Device switching**: Uses the undocumented `IPolicyConfig` COM interface (`SetDefaultEndpoint`) to change the system default device. Falls back to PowerShell `Set-AudioDevice` if available.
- **Change notifications**: Registers an `IMMNotificationClient` callback for real-time device change events, with a 2-second polling fallback.
- **Always-on-top**: Win32 `SetWindowPos` with `HWND_TOPMOST` for true taskbar-level always-on-top, re-asserted every 500ms.
- **Bluetooth management**: Calls `BluetoothSetServiceState` via `ctypes` to toggle A2DP Sink and HFP service connections. Connect/disconnect runs on a background `QThread` to keep the UI responsive (~5–7 seconds per operation).
- **BT endpoint reconciliation**: Windows assigns a new audio endpoint ID each time a Bluetooth device reconnects. Audioflip detects this by matching the BT device name inside the endpoint friendly name and automatically migrates the favourite, metadata, and icon overrides to the new ID.
- **UI**: PyQt6 frameless window with 10 colour themes, drag support, popup dropdown, and system tray icon.
