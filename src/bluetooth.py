"""Windows Bluetooth API wrappers for connecting/disconnecting audio devices.

Primary approach: Win32 BluetoothSetServiceState via BluetoothApis.dll
with an explicit radio handle (required by some Intel/Dell BT stacks).

Fallback: PowerShell Disable-PnpDevice / Enable-PnpDevice when the
Win32 API fails (works across all BT controller vendors).

Gracefully degrades if Bluetooth hardware is unavailable.
"""

from __future__ import annotations

import atexit
import ctypes
import ctypes.wintypes
import json
import logging
import subprocess
import time
from ctypes import (
    POINTER,
    byref,
    c_ubyte,
    c_ulonglong,
    c_void_p,
    memmove,
    sizeof,
)
from ctypes.wintypes import BOOL, DWORD, ULONG, WORD

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Bluetooth service GUIDs (A2DP Sink + Hands-Free Profile)
# ---------------------------------------------------------------------------
class _GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]

def _make_guid(d1: int, d2: int, d3: int, d4: bytes) -> _GUID:
    g = _GUID()
    g.Data1 = d1
    g.Data2 = d2
    g.Data3 = d3
    g.Data4 = (ctypes.c_ubyte * 8)(*d4)
    return g

# A2DP Sink: {0000110b-0000-1000-8000-00805f9b34fb}
_A2DP_SINK = _make_guid(0x0000110B, 0x0000, 0x1000, b"\x80\x00\x00\x80\x5f\x9b\x34\xfb")

# Hands-Free Profile (HFP): {0000111e-0000-1000-8000-00805f9b34fb}
_HFP = _make_guid(0x0000111E, 0x0000, 0x1000, b"\x80\x00\x00\x80\x5f\x9b\x34\xfb")

# Audio Source (A2DP Source): {0000110a-0000-1000-8000-00805f9b34fb}
_A2DP_SOURCE = _make_guid(0x0000110A, 0x0000, 0x1000, b"\x80\x00\x00\x80\x5f\x9b\x34\xfb")

_AUDIO_SERVICE_GUIDS = [_A2DP_SINK, _HFP, _A2DP_SOURCE]

# ---------------------------------------------------------------------------
# Bluetooth structures
# ---------------------------------------------------------------------------
class SYSTEMTIME(ctypes.Structure):
    _fields_ = [
        ("wYear", WORD),
        ("wMonth", WORD),
        ("wDayOfWeek", WORD),
        ("wDay", WORD),
        ("wHour", WORD),
        ("wMinute", WORD),
        ("wSecond", WORD),
        ("wMilliseconds", WORD),
    ]


class BLUETOOTH_ADDRESS(ctypes.Union):
    _fields_ = [
        ("ullLong", c_ulonglong),
        ("rgBytes", c_ubyte * 6),
    ]


class BLUETOOTH_DEVICE_INFO(ctypes.Structure):
    _fields_ = [
        ("dwSize", DWORD),
        ("Address", BLUETOOTH_ADDRESS),
        ("ulClassofDevice", ULONG),
        ("fConnected", BOOL),
        ("fRemembered", BOOL),
        ("fAuthenticated", BOOL),
        ("stLastSeen", SYSTEMTIME),
        ("stLastUsed", SYSTEMTIME),
        ("szName", ctypes.c_wchar * 248),
    ]


class BLUETOOTH_DEVICE_SEARCH_PARAMS(ctypes.Structure):
    _fields_ = [
        ("dwSize", DWORD),
        ("fReturnAuthenticated", BOOL),
        ("fReturnRemembered", BOOL),
        ("fReturnUnknown", BOOL),
        ("fReturnConnected", BOOL),
        ("fIssueInquiry", BOOL),
        ("cTimeoutMultiplier", c_ubyte),
        ("hRadio", c_void_p),  # HANDLE — use c_void_p for 64-bit safety
    ]


class BLUETOOTH_FIND_RADIO_PARAMS(ctypes.Structure):
    _fields_ = [
        ("dwSize", DWORD),
    ]


# ---------------------------------------------------------------------------
# Load BluetoothApis.dll with proper function prototypes
# ---------------------------------------------------------------------------
_bt: ctypes.WinDLL | None = None
_bt_available: bool | None = None
_bt_auth_func = None  # BluetoothAuthenticateDeviceEx (may be in bthprops.cpl)
_radio_handle: c_void_p | None = None


def _load_bt() -> ctypes.WinDLL | None:
    """Load BluetoothApis.dll and set up function prototypes."""
    global _bt, _bt_available
    if _bt_available is not None:
        return _bt

    try:
        dll = ctypes.WinDLL("BluetoothApis.dll", use_last_error=True)
    except OSError:
        try:
            dll = ctypes.WinDLL("bthprops.cpl", use_last_error=True)
        except OSError:
            _bt = None
            _bt_available = False
            log.info("Bluetooth APIs not available on this system")
            return None

    # --- Set up function prototypes (critical for 64-bit) ---

    # BluetoothFindFirstRadio / BluetoothFindRadioClose
    dll.BluetoothFindFirstRadio.argtypes = [
        POINTER(BLUETOOTH_FIND_RADIO_PARAMS),
        POINTER(c_void_p),
    ]
    dll.BluetoothFindFirstRadio.restype = c_void_p  # HBLUETOOTH_RADIO_FIND

    dll.BluetoothFindRadioClose.argtypes = [c_void_p]
    dll.BluetoothFindRadioClose.restype = BOOL

    # BluetoothFindFirstDevice / BluetoothFindNextDevice / Close
    dll.BluetoothFindFirstDevice.argtypes = [
        POINTER(BLUETOOTH_DEVICE_SEARCH_PARAMS),
        POINTER(BLUETOOTH_DEVICE_INFO),
    ]
    dll.BluetoothFindFirstDevice.restype = c_void_p  # HANDLE

    dll.BluetoothFindNextDevice.argtypes = [c_void_p, POINTER(BLUETOOTH_DEVICE_INFO)]
    dll.BluetoothFindNextDevice.restype = BOOL

    dll.BluetoothFindDeviceClose.argtypes = [c_void_p]
    dll.BluetoothFindDeviceClose.restype = BOOL

    # BluetoothSetServiceState
    dll.BluetoothSetServiceState.argtypes = [
        c_void_p,                       # hRadio
        POINTER(BLUETOOTH_DEVICE_INFO),
        POINTER(_GUID),
        DWORD,
    ]
    dll.BluetoothSetServiceState.restype = DWORD

    # BluetoothAuthenticateDeviceEx (SSP pairing for BT 2.1+)
    # This function lives in bthprops.cpl on some Windows builds and may not
    # be exported from BluetoothApis.dll.  Try the main dll first, then fall
    # back to bthprops.cpl, and finally give up gracefully so core BT
    # connect/disconnect still works even when pairing is unavailable.
    global _bt_auth_func
    _bt_auth_func = None
    for _src in (dll, "bthprops.cpl"):
        try:
            _lib = _src if _src is dll else ctypes.WinDLL(_src)
            _lib.BluetoothAuthenticateDeviceEx.argtypes = [
                c_void_p,                       # hwndParentIn (NULL ok)
                c_void_p,                       # hRadio
                POINTER(BLUETOOTH_DEVICE_INFO), # pbtdiInout
                c_void_p,                       # pbtOobData (NULL for SSP)
                DWORD,                          # authenticationRequirement
            ]
            _lib.BluetoothAuthenticateDeviceEx.restype = DWORD
            _bt_auth_func = _lib.BluetoothAuthenticateDeviceEx
            log.info("BluetoothAuthenticateDeviceEx loaded from %s",
                     "BluetoothApis.dll" if _src is dll else _src)
            break
        except (AttributeError, OSError):
            continue
    if _bt_auth_func is None:
        log.warning("BluetoothAuthenticateDeviceEx not available — "
                    "new-device pairing will be disabled")

    _bt = dll
    _bt_available = True
    log.info("Bluetooth APIs loaded successfully")
    return _bt


def _get_radio_handle() -> c_void_p | None:
    """Get a handle to the first Bluetooth radio on the system.

    Some BT stacks (Intel/Dell) require an explicit radio handle instead
    of NULL. The handle is cached for the lifetime of the process.
    """
    global _radio_handle
    if _radio_handle is not None:
        return _radio_handle

    dll = _load_bt()
    if dll is None:
        return None

    params = BLUETOOTH_FIND_RADIO_PARAMS()
    params.dwSize = sizeof(BLUETOOTH_FIND_RADIO_PARAMS)
    h_radio = c_void_p()

    try:
        h_find = dll.BluetoothFindFirstRadio(byref(params), byref(h_radio))
    except Exception as exc:
        log.debug("BluetoothFindFirstRadio exception: %s", exc)
        return None

    if not h_find:
        log.debug("BluetoothFindFirstRadio returned NULL (no radios found)")
        return None

    # Close the *find* handle but keep the *radio* handle open
    try:
        dll.BluetoothFindRadioClose(h_find)
    except Exception:
        pass

    _radio_handle = h_radio
    log.info("Obtained Bluetooth radio handle: %s", h_radio.value)
    atexit.register(_close_radio_handle)
    return _radio_handle


def _close_radio_handle() -> None:
    """Close the cached radio handle on process exit."""
    global _radio_handle
    if _radio_handle is not None:
        try:
            ctypes.windll.kernel32.CloseHandle(_radio_handle)
        except Exception:
            pass
        _radio_handle = None


def is_bluetooth_available() -> bool:
    """Return True if Windows Bluetooth APIs are accessible."""
    _load_bt()
    return _bt_available is True


# ---------------------------------------------------------------------------
# Paired device name cache (used by audio_manager for BT detection fallback)
# ---------------------------------------------------------------------------
_paired_names_cache: list[str] | None = None
_paired_names_cache_time: float = 0.0
_PAIRED_NAMES_CACHE_TTL = 30.0  # seconds


def get_paired_device_names() -> list[str]:
    """Return names of all paired Bluetooth devices (cached).

    Results are cached for 30 seconds to avoid calling Win32 BT
    enumeration on every audio device poll (~1/s).  Used by
    AudioManager.enumerate_devices() as a fallback when the Windows
    audio enumerator property doesn't indicate Bluetooth (e.g. Intel
    audio controllers that proxy BT audio as INTELAUDIO).
    """
    global _paired_names_cache, _paired_names_cache_time
    now = time.monotonic()
    if _paired_names_cache is not None and (now - _paired_names_cache_time) < _PAIRED_NAMES_CACHE_TTL:
        return _paired_names_cache

    devices = _find_paired_devices()
    _paired_names_cache = [d.szName.strip() for d in devices if d.szName.strip()]
    _paired_names_cache_time = now
    log.debug("Refreshed paired BT device names cache: %s", _paired_names_cache)
    return _paired_names_cache


# ---------------------------------------------------------------------------
# Win32 API helpers
# ---------------------------------------------------------------------------
def _find_paired_devices() -> list[BLUETOOTH_DEVICE_INFO]:
    """Enumerate all paired/remembered Bluetooth devices."""
    dll = _load_bt()
    if dll is None:
        return []

    radio = _get_radio_handle()

    params = BLUETOOTH_DEVICE_SEARCH_PARAMS()
    params.dwSize = sizeof(BLUETOOTH_DEVICE_SEARCH_PARAMS)
    params.fReturnAuthenticated = True
    params.fReturnRemembered = True
    params.fReturnUnknown = False
    params.fReturnConnected = True
    params.fIssueInquiry = False
    params.cTimeoutMultiplier = 0
    params.hRadio = radio  # explicit handle (or None → default)

    device_info = BLUETOOTH_DEVICE_INFO()
    device_info.dwSize = sizeof(BLUETOOTH_DEVICE_INFO)

    devices: list[BLUETOOTH_DEVICE_INFO] = []

    try:
        h_find = dll.BluetoothFindFirstDevice(byref(params), byref(device_info))
    except Exception as exc:
        log.warning("BluetoothFindFirstDevice exception: %s", exc)
        return []

    if not h_find:
        err = ctypes.get_last_error()
        log.debug("BluetoothFindFirstDevice returned NULL (last error: %d)", err)
        return []

    log.debug("Found paired BT device: '%s' (connected=%s)", device_info.szName, bool(device_info.fConnected))
    devices.append(_copy_device_info(device_info))

    while True:
        next_info = BLUETOOTH_DEVICE_INFO()
        next_info.dwSize = sizeof(BLUETOOTH_DEVICE_INFO)
        try:
            found = dll.BluetoothFindNextDevice(h_find, byref(next_info))
        except Exception:
            break
        if not found:
            break
        log.debug("Found paired BT device: '%s' (connected=%s)", next_info.szName, bool(next_info.fConnected))
        devices.append(_copy_device_info(next_info))

    try:
        dll.BluetoothFindDeviceClose(h_find)
    except Exception:
        pass

    log.info("Enumerated %d paired Bluetooth devices", len(devices))
    return devices


def _copy_device_info(src: BLUETOOTH_DEVICE_INFO) -> BLUETOOTH_DEVICE_INFO:
    """Make a deep copy of a BLUETOOTH_DEVICE_INFO structure."""
    dst = BLUETOOTH_DEVICE_INFO()
    memmove(byref(dst), byref(src), sizeof(BLUETOOTH_DEVICE_INFO))
    return dst


# ---------------------------------------------------------------------------
# Discovery & Pairing (for "Scan for Bluetooth" feature)
# ---------------------------------------------------------------------------
_AUDIO_VIDEO_MAJOR_CLASS = 0x04  # Bluetooth CoD major class: Audio/Video


def discover_audio_devices(timeout_multiplier: int = 4) -> list[dict]:
    """Scan for nearby Bluetooth audio devices (blocking ~5-12s).

    Issues an active inquiry scan and filters results to Audio/Video
    major class (0x04).  Returns a list of dicts:

        {name, address, is_paired, _device_info}

    ``is_paired`` is True if the device is already authenticated or
    remembered.  ``_device_info`` holds the raw BLUETOOTH_DEVICE_INFO
    struct for passing to ``pair_and_connect_device()``.
    """
    dll = _load_bt()
    if dll is None:
        log.warning("discover_audio_devices: BT APIs not available")
        return []

    radio = _get_radio_handle()

    params = BLUETOOTH_DEVICE_SEARCH_PARAMS()
    params.dwSize = sizeof(BLUETOOTH_DEVICE_SEARCH_PARAMS)
    params.fReturnAuthenticated = True
    params.fReturnRemembered = True
    params.fReturnUnknown = True          # discover NEW devices
    params.fReturnConnected = True
    params.fIssueInquiry = True           # active scan
    params.cTimeoutMultiplier = timeout_multiplier  # ~1.28s per unit
    params.hRadio = radio

    device_info = BLUETOOTH_DEVICE_INFO()
    device_info.dwSize = sizeof(BLUETOOTH_DEVICE_INFO)

    raw_devices: list[BLUETOOTH_DEVICE_INFO] = []

    try:
        h_find = dll.BluetoothFindFirstDevice(byref(params), byref(device_info))
    except Exception as exc:
        log.warning("discover_audio_devices: BluetoothFindFirstDevice exception: %s", exc)
        return []

    if not h_find:
        err = ctypes.get_last_error()
        log.info("discover_audio_devices: no devices found (last error: %d)", err)
        return []

    raw_devices.append(_copy_device_info(device_info))

    while True:
        next_info = BLUETOOTH_DEVICE_INFO()
        next_info.dwSize = sizeof(BLUETOOTH_DEVICE_INFO)
        try:
            found = dll.BluetoothFindNextDevice(h_find, byref(next_info))
        except Exception:
            break
        if not found:
            break
        raw_devices.append(_copy_device_info(next_info))

    try:
        dll.BluetoothFindDeviceClose(h_find)
    except Exception:
        pass

    log.info("discover_audio_devices: found %d total devices", len(raw_devices))

    # Filter to Audio/Video major class
    results: list[dict] = []
    for dev in raw_devices:
        major = (dev.ulClassofDevice >> 8) & 0x1F
        name = dev.szName.strip()
        if major != _AUDIO_VIDEO_MAJOR_CLASS:
            log.debug("  skipping '%s' (major class 0x%02X, not audio)", name, major)
            continue
        addr = ":".join(f"{b:02X}" for b in reversed(dev.Address.rgBytes[:6]))
        is_paired = bool(dev.fAuthenticated) or bool(dev.fRemembered)
        log.info("  audio device: '%s' addr=%s paired=%s connected=%s",
                 name, addr, is_paired, bool(dev.fConnected))
        results.append({
            "name": name,
            "address": addr,
            "is_paired": is_paired,
            "_device_info": dev,
        })

    log.info("discover_audio_devices: %d audio devices after filtering", len(results))
    return results


def pair_and_connect_device(device_info: BLUETOOTH_DEVICE_INFO) -> bool:
    """Pair a discovered BT device and enable audio services (A2DP + HFP).

    Uses BluetoothAuthenticateDeviceEx for Secure Simple Pairing.
    After pairing, enables all audio service GUIDs via
    BluetoothSetServiceState.

    Returns True if pairing succeeded and at least one audio service
    was enabled.
    """
    dll = _load_bt()
    if dll is None:
        return False

    radio = _get_radio_handle()
    dev_name = device_info.szName.strip()

    # Already authenticated?
    if not device_info.fAuthenticated:
        if _bt_auth_func is None:
            log.error("Cannot pair '%s' — BluetoothAuthenticateDeviceEx "
                      "not available on this system", dev_name)
            return False
        # MITMProtectionNotRequiredGeneralBonding = 0x04
        log.info("Pairing '%s' via BluetoothAuthenticateDeviceEx …", dev_name)
        try:
            result = _bt_auth_func(
                None,                   # hwndParentIn
                radio,                  # hRadio
                byref(device_info),     # pbtdiInout
                None,                   # pbtOobData (NULL → SSP)
                DWORD(0x04),            # authenticationRequirement
            )
        except Exception as exc:
            log.error("BluetoothAuthenticateDeviceEx exception for '%s': %s", dev_name, exc)
            return False

        if result != 0:
            log.warning("BluetoothAuthenticateDeviceEx returned %d (0x%X) for '%s'",
                        result, result, dev_name)
            return False
        log.info("Pairing succeeded for '%s'", dev_name)
    else:
        log.info("'%s' is already authenticated, skipping pairing", dev_name)

    # Enable audio services
    success = False
    for guid in _AUDIO_SERVICE_GUIDS:
        if _set_service_state(device_info, guid, enable=True):
            success = True

    # Invalidate paired names cache so next poll picks up the new device
    global _paired_names_cache
    _paired_names_cache = None

    if success:
        log.info("Audio services enabled for '%s'", dev_name)
    else:
        log.warning("No audio services could be enabled for '%s'", dev_name)

    return success


def _match_device_by_name(name: str) -> BLUETOOTH_DEVICE_INFO | None:
    """Find a paired BT device whose name matches the audio endpoint.

    The Windows audio endpoint name often contains the BT device name
    plus extra text (e.g. "JBL TUNE510BT Stereo"), so we do a
    case-insensitive substring match.
    """
    name_lower = name.lower()
    paired = _find_paired_devices()
    log.info("Trying to match audio name '%s' against %d paired BT devices", name, len(paired))

    for dev in paired:
        bt_name = dev.szName.strip()
        if not bt_name:
            continue
        bt_lower = bt_name.lower()
        # BT device name appears inside the audio endpoint name
        if bt_lower in name_lower:
            log.info("Matched BT device '%s' (substring of audio name '%s')", bt_name, name)
            return dev
        # Or audio name appears inside BT device name
        if name_lower in bt_lower:
            log.info("Matched BT device '%s' (audio name '%s' is substring)", bt_name, name)
            return dev

    log.warning("No BT device matched audio name '%s'. Paired devices: %s",
                name, [d.szName.strip() for d in paired])
    return None


def _set_service_state(
    device_info: BLUETOOTH_DEVICE_INFO, service_guid: _GUID, enable: bool
) -> bool:
    """Enable or disable a Bluetooth service on a device."""
    dll = _load_bt()
    if dll is None:
        return False

    radio = _get_radio_handle()
    # BLUETOOTH_SERVICE_DISABLE = 0x00, BLUETOOTH_SERVICE_ENABLE = 0x01
    flags = 0x01 if enable else 0x00
    action = "enable" if enable else "disable"

    try:
        result = dll.BluetoothSetServiceState(
            radio,                 # explicit radio handle (or None)
            byref(device_info),
            byref(service_guid),
            flags,
        )
        if result != 0:
            log.warning(
                "BluetoothSetServiceState(%s) returned error %d (0x%X) for '%s'",
                action, result, result, device_info.szName,
            )
            return False
        log.info("BluetoothSetServiceState(%s) succeeded for '%s'", action, device_info.szName)
        return True
    except Exception as exc:
        log.error("BluetoothSetServiceState(%s) exception for '%s': %s", action, device_info.szName, exc)
        return False


def _win32_connect(device_name: str) -> bool:
    """Connect via Win32 BluetoothSetServiceState."""
    dev = _match_device_by_name(device_name)
    if dev is None:
        return False

    log.info("Win32: Connecting BT device '%s' (matched from '%s')", dev.szName, device_name)
    success = False
    for guid in _AUDIO_SERVICE_GUIDS:
        if _set_service_state(dev, guid, enable=True):
            success = True
    return success


def _win32_disconnect(device_name: str) -> bool:
    """Disconnect via Win32 BluetoothSetServiceState."""
    dev = _match_device_by_name(device_name)
    if dev is None:
        return False

    log.info("Win32: Disconnecting BT device '%s' (matched from '%s')", dev.szName, device_name)
    success = False
    for guid in _AUDIO_SERVICE_GUIDS:
        if _set_service_state(dev, guid, enable=False):
            success = True
    return success


# ---------------------------------------------------------------------------
# PowerShell PnP fallback
# ---------------------------------------------------------------------------
def _find_bt_pnp_instance_id(device_name: str) -> str | None:
    """Find the PnP instance ID for a Bluetooth device by name.

    Uses Get-PnpDevice to list Bluetooth devices and matches by
    friendly name (case-insensitive substring).
    """
    name_lower = device_name.lower()
    cmd = (
        'powershell -NoProfile -Command "'
        "Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue "
        "| Select-Object InstanceId, FriendlyName, Status "
        '| ConvertTo-Json -Compress"'
    )
    try:
        result = subprocess.run(
            cmd, shell=True, capture_output=True, text=True, timeout=10,
        )
        if result.returncode != 0:
            log.debug("Get-PnpDevice failed: %s", result.stderr.strip())
            return None
    except Exception as exc:
        log.debug("Get-PnpDevice exception: %s", exc)
        return None

    try:
        data = json.loads(result.stdout)
    except (json.JSONDecodeError, ValueError):
        log.debug("Get-PnpDevice returned non-JSON: %s", result.stdout[:200])
        return None

    # PowerShell returns a single object (not list) when there's only one device
    if isinstance(data, dict):
        data = [data]

    for dev in data:
        friendly = (dev.get("FriendlyName") or "").lower()
        if not friendly:
            continue
        if friendly in name_lower or name_lower in friendly:
            instance_id = dev.get("InstanceId", "")
            log.info("PowerShell: Matched PnP device '%s' (InstanceId: %s)",
                     dev.get("FriendlyName"), instance_id)
            return instance_id

    log.debug("PowerShell: No PnP device matched '%s'. Devices: %s",
              device_name, [d.get("FriendlyName") for d in data])
    return None


def _powershell_enable(instance_id: str) -> bool:
    """Enable a PnP device (reconnect) via PowerShell."""
    cmd = (
        f'powershell -NoProfile -Command "'
        f"Enable-PnpDevice -InstanceId '{instance_id}' -Confirm:\\$false"
        f' -ErrorAction Stop"'
    )
    try:
        result = subprocess.run(
            cmd, shell=True, capture_output=True, text=True, timeout=15,
        )
        if result.returncode == 0:
            log.info("PowerShell: Enable-PnpDevice succeeded for %s", instance_id)
            return True
        stderr = result.stderr.strip()
        if "Access" in stderr or "denied" in stderr or "administrator" in stderr.lower():
            log.warning("PowerShell: Enable-PnpDevice needs admin privileges")
        else:
            log.warning("PowerShell: Enable-PnpDevice failed: %s", stderr)
        return False
    except Exception as exc:
        log.warning("PowerShell: Enable-PnpDevice exception: %s", exc)
        return False


def _powershell_disable(instance_id: str) -> bool:
    """Disable a PnP device (disconnect) via PowerShell."""
    cmd = (
        f'powershell -NoProfile -Command "'
        f"Disable-PnpDevice -InstanceId '{instance_id}' -Confirm:\\$false"
        f' -ErrorAction Stop"'
    )
    try:
        result = subprocess.run(
            cmd, shell=True, capture_output=True, text=True, timeout=15,
        )
        if result.returncode == 0:
            log.info("PowerShell: Disable-PnpDevice succeeded for %s", instance_id)
            return True
        stderr = result.stderr.strip()
        if "Access" in stderr or "denied" in stderr or "administrator" in stderr.lower():
            log.warning("PowerShell: Disable-PnpDevice needs admin privileges")
        else:
            log.warning("PowerShell: Disable-PnpDevice failed: %s", stderr)
        return False
    except Exception as exc:
        log.warning("PowerShell: Disable-PnpDevice exception: %s", exc)
        return False


def _powershell_connect(device_name: str) -> bool:
    """Fallback: connect a BT device via PowerShell Enable-PnpDevice."""
    instance_id = _find_bt_pnp_instance_id(device_name)
    if not instance_id:
        log.info("PowerShell fallback: no PnP device found for '%s'", device_name)
        return False
    return _powershell_enable(instance_id)


def _powershell_disconnect(device_name: str) -> bool:
    """Fallback: disconnect a BT device via PowerShell Disable-PnpDevice."""
    instance_id = _find_bt_pnp_instance_id(device_name)
    if not instance_id:
        log.info("PowerShell fallback: no PnP device found for '%s'", device_name)
        return False
    return _powershell_disable(instance_id)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def bluetooth_connect(device_name: str) -> bool:
    """Attempt to connect a Bluetooth audio device by name.

    1. Try Win32 BluetoothSetServiceState (with explicit radio handle).
    2. If that fails, fall back to PowerShell Enable-PnpDevice.

    Returns True if either method succeeded.
    """
    if not is_bluetooth_available():
        log.warning("Bluetooth not available, cannot connect '%s'", device_name)
        # Still try PowerShell — the Win32 DLL might be missing but
        # PowerShell PnP can still work
        log.info("Attempting PowerShell fallback for connect '%s'", device_name)
        return _powershell_connect(device_name)

    if _win32_connect(device_name):
        log.info("BT connect via Win32 API succeeded for '%s'", device_name)
        return True

    log.info("Win32 API failed for '%s', trying PowerShell fallback", device_name)
    if _powershell_connect(device_name):
        log.info("BT connect via PowerShell succeeded for '%s'", device_name)
        return True

    log.warning("All BT connect methods failed for '%s'", device_name)
    return False


def bluetooth_disconnect(device_name: str) -> bool:
    """Attempt to disconnect a Bluetooth audio device by name.

    1. Try Win32 BluetoothSetServiceState (with explicit radio handle).
    2. If that fails, fall back to PowerShell Disable-PnpDevice.

    Returns True if either method succeeded.
    """
    if not is_bluetooth_available():
        log.warning("Bluetooth not available, cannot disconnect '%s'", device_name)
        log.info("Attempting PowerShell fallback for disconnect '%s'", device_name)
        return _powershell_disconnect(device_name)

    if _win32_disconnect(device_name):
        log.info("BT disconnect via Win32 API succeeded for '%s'", device_name)
        return True

    log.info("Win32 API failed for '%s', trying PowerShell fallback", device_name)
    if _powershell_disconnect(device_name):
        log.info("BT disconnect via PowerShell succeeded for '%s'", device_name)
        return True

    log.warning("All BT disconnect methods failed for '%s'", device_name)
    return False
