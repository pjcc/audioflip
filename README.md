# audioflip

A persistent, always-on-top Windows desktop widget for switching audio devices.

## Features

- Compact widget pinned to your screen showing the current audio device
- Click to open a dropdown of all input/output devices
- Switch the system default device with one click
- Auto-assigned icons based on device name (headphones, speakers, Bluetooth, etc.)
- Customisable per-device icons via right-click menu
- Favourite devices (right-click a device in the dropdown)
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

```bash
# Clone or download the project, then:
cd audioflip

# Create a virtual environment (recommended)
python -m venv .venv
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Running

```bash
python run.py
```

## Building a Standalone .exe

```bash
# Install PyInstaller (included in requirements.txt)
pyinstaller audioflip.spec
```

The output will be at `dist/audioflip.exe` — a single portable executable.

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
    ├── icons.py            # Icon matching and rendering
    └── ui.py               # PyQt6 widget, dropdown, context menu
```

## How It Works

- **Device enumeration**: Uses `pycaw` and `comtypes` to query `IMMDeviceEnumerator` for active audio endpoints.
- **Device switching**: Uses the undocumented `IPolicyConfig` COM interface (`SetDefaultEndpoint`) to change the system default device. Falls back to PowerShell `Set-AudioDevice` if available.
- **Change notifications**: Registers an `IMMNotificationClient` callback for real-time device change events, with a 2-second polling fallback.
- **Always-on-top**: Win32 `SetWindowPos` with `HWND_TOPMOST` for true taskbar-level always-on-top, re-asserted every 500ms.
- **UI**: PyQt6 frameless window with 10 colour themes, drag support, popup dropdown, and system tray icon.
