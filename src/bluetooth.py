"""Windows Bluetooth API wrappers for connecting/disconnecting audio devices.

Uses ctypes to call BluetoothApis.dll functions:
- Enumerate paired Bluetooth devices
- Match a BT device to an audio endpoint by name substring
- Connect/disconnect via BluetoothSetServiceState with A2DP + HFP service GUIDs

Gracefully degrades if Bluetooth hardware is unavailable.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import logging
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


# ---------------------------------------------------------------------------
# Load BluetoothApis.dll with proper function prototypes
# ---------------------------------------------------------------------------
_bt: ctypes.WinDLL | None = None
_bt_available: bool | None = None


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

    # HBLUETOOTH_DEVICE_FIND BluetoothFindFirstDevice(
    #     const BLUETOOTH_DEVICE_SEARCH_PARAMS *pbtsp,
    #     BLUETOOTH_DEVICE_INFO *pbtdi)
    dll.BluetoothFindFirstDevice.argtypes = [
        POINTER(BLUETOOTH_DEVICE_SEARCH_PARAMS),
        POINTER(BLUETOOTH_DEVICE_INFO),
    ]
    dll.BluetoothFindFirstDevice.restype = c_void_p  # HANDLE

    # BOOL BluetoothFindNextDevice(
    #     HBLUETOOTH_DEVICE_FIND hFind,
    #     BLUETOOTH_DEVICE_INFO *pbtdi)
    dll.BluetoothFindNextDevice.argtypes = [c_void_p, POINTER(BLUETOOTH_DEVICE_INFO)]
    dll.BluetoothFindNextDevice.restype = BOOL

    # BOOL BluetoothFindDeviceClose(HBLUETOOTH_DEVICE_FIND hFind)
    dll.BluetoothFindDeviceClose.argtypes = [c_void_p]
    dll.BluetoothFindDeviceClose.restype = BOOL

    # DWORD BluetoothSetServiceState(
    #     HANDLE hRadio,
    #     const BLUETOOTH_DEVICE_INFO *pbtdi,
    #     const GUID *pGuidService,
    #     DWORD dwServiceFlags)
    dll.BluetoothSetServiceState.argtypes = [
        c_void_p,                       # hRadio (NULL for default)
        POINTER(BLUETOOTH_DEVICE_INFO),
        POINTER(_GUID),
        DWORD,
    ]
    dll.BluetoothSetServiceState.restype = DWORD

    _bt = dll
    _bt_available = True
    log.info("Bluetooth APIs loaded successfully")
    return _bt


def is_bluetooth_available() -> bool:
    """Return True if Windows Bluetooth APIs are accessible."""
    _load_bt()
    return _bt_available is True


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------
def _find_paired_devices() -> list[BLUETOOTH_DEVICE_INFO]:
    """Enumerate all paired/remembered Bluetooth devices."""
    dll = _load_bt()
    if dll is None:
        return []

    params = BLUETOOTH_DEVICE_SEARCH_PARAMS()
    params.dwSize = sizeof(BLUETOOTH_DEVICE_SEARCH_PARAMS)
    params.fReturnAuthenticated = True
    params.fReturnRemembered = True
    params.fReturnUnknown = False
    params.fReturnConnected = True
    params.fIssueInquiry = False
    params.cTimeoutMultiplier = 0
    params.hRadio = None  # use default radio

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

    # BLUETOOTH_SERVICE_DISABLE = 0x00, BLUETOOTH_SERVICE_ENABLE = 0x01
    flags = 0x01 if enable else 0x00
    action = "enable" if enable else "disable"

    try:
        result = dll.BluetoothSetServiceState(
            None,                  # hRadio — NULL uses default radio
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


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def bluetooth_connect(device_name: str) -> bool:
    """Attempt to connect a Bluetooth audio device by name.

    Matches the audio endpoint name against paired BT devices, then
    enables A2DP and HFP services.

    Returns True if at least one service was enabled successfully.
    """
    if not is_bluetooth_available():
        log.warning("Bluetooth not available, cannot connect '%s'", device_name)
        return False

    dev = _match_device_by_name(device_name)
    if dev is None:
        return False

    log.info("Connecting BT device: '%s' (matched from audio name '%s')", dev.szName, device_name)
    success = False
    for guid in _AUDIO_SERVICE_GUIDS:
        if _set_service_state(dev, guid, enable=True):
            success = True
    log.info("BT connect result for '%s': %s", device_name, "success" if success else "FAILED")
    return success


def bluetooth_disconnect(device_name: str) -> bool:
    """Attempt to disconnect a Bluetooth audio device by name.

    Matches the audio endpoint name against paired BT devices, then
    disables A2DP and HFP services.

    Returns True if at least one service was disabled successfully.
    """
    if not is_bluetooth_available():
        log.warning("Bluetooth not available, cannot disconnect '%s'", device_name)
        return False

    dev = _match_device_by_name(device_name)
    if dev is None:
        return False

    log.info("Disconnecting BT device: '%s' (matched from audio name '%s')", dev.szName, device_name)
    success = False
    for guid in _AUDIO_SERVICE_GUIDS:
        if _set_service_state(dev, guid, enable=False):
            success = True
    log.info("BT disconnect result for '%s': %s", device_name, "success" if success else "FAILED")
    return success
