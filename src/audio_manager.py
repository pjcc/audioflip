"""Audio device enumeration and switching via Windows Core Audio APIs.

Uses pycaw / comtypes to enumerate devices and the undocumented
IPolicyConfig COM interface to set the default audio device.
Provides event-driven notifications for device changes.
"""

from __future__ import annotations

import enum
import logging
from dataclasses import dataclass
from typing import Callable

import comtypes
from comtypes import GUID, HRESULT, COMMETHOD
from ctypes import POINTER, c_uint, c_wchar_p, cast
from ctypes.wintypes import LPCWSTR, DWORD, BOOL

from pycaw.pycaw import (
    AudioUtilities,
    IAudioEndpointVolume,
    IMMDeviceEnumerator,
    IMMDevice,
    IMMNotificationClient,
    EDataFlow,
    ERole,
    DEVICE_STATE,
)

log = logging.getLogger(__name__)

_CLSCTX_ALL = 0x17  # CLSCTX_INPROC_SERVER | INPROC_HANDLER | LOCAL_SERVER | REMOTE_SERVER

# ---------------------------------------------------------------------------
# Data types
# ---------------------------------------------------------------------------
class DeviceFlow(enum.Enum):
    OUTPUT = "output"
    INPUT = "input"


@dataclass(frozen=True)
class AudioDevice:
    """Represents an audio endpoint device."""
    id: str
    name: str
    flow: DeviceFlow
    is_default: bool = False


# ---------------------------------------------------------------------------
# IPolicyConfig COM interface (undocumented, Windows Vista+)
# ---------------------------------------------------------------------------
class IPolicyConfig(comtypes.IUnknown):
    _iid_ = GUID("{f8679f50-850a-41cf-9c72-430f290290c8}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetMixFormat",
                  (["in"], LPCWSTR, "pszDeviceName"),
                  (["out"], POINTER(c_uint), "ppFormat")),
        COMMETHOD([], HRESULT, "GetDeviceFormat",
                  (["in"], LPCWSTR, "pszDeviceName"),
                  (["in"], BOOL, "bDefault"),
                  (["out"], POINTER(c_uint), "ppFormat")),
        COMMETHOD([], HRESULT, "ResetDeviceFormat",
                  (["in"], LPCWSTR, "pszDeviceName")),
        COMMETHOD([], HRESULT, "SetDeviceFormat",
                  (["in"], LPCWSTR, "pszDeviceName"),
                  (["in"], c_uint, "pEndpointFormat"),
                  (["in"], c_uint, "mixFormat")),
        COMMETHOD([], HRESULT, "GetProcessingPeriod",
                  (["in"], LPCWSTR, "pszDeviceName"),
                  (["in"], BOOL, "bDefault"),
                  (["out"], POINTER(c_uint), "pmftDefaultPeriod"),
                  (["out"], POINTER(c_uint), "pmftMinimumPeriod")),
        COMMETHOD([], HRESULT, "SetProcessingPeriod",
                  (["in"], LPCWSTR, "pszDeviceName"),
                  (["in"], c_uint, "pmftPeriod")),
        COMMETHOD([], HRESULT, "GetShareMode",
                  (["in"], LPCWSTR, "pszDeviceName"),
                  (["out"], POINTER(c_uint), "pMode")),
        COMMETHOD([], HRESULT, "SetShareMode",
                  (["in"], LPCWSTR, "pszDeviceName"),
                  (["in"], c_uint, "mode")),
        COMMETHOD([], HRESULT, "GetPropertyValue",
                  (["in"], LPCWSTR, "pszDeviceName"),
                  (["in"], c_uint, "key"),
                  (["out"], POINTER(c_uint), "pv")),
        COMMETHOD([], HRESULT, "SetPropertyValue",
                  (["in"], LPCWSTR, "pszDeviceName"),
                  (["in"], c_uint, "key"),
                  (["in"], c_uint, "pv")),
        COMMETHOD([], HRESULT, "SetDefaultEndpoint",
                  (["in"], LPCWSTR, "pszDeviceName"),
                  (["in"], c_uint, "eRole")),
        COMMETHOD([], HRESULT, "SetEndpointVisibility",
                  (["in"], LPCWSTR, "pszDeviceName"),
                  (["in"], BOOL, "bVisible")),
    ]


class CPolicyConfigClient(comtypes.CoClass):
    _reg_clsid_ = GUID("{870af99c-171d-4f9e-af0d-e63df40c2bc9}")
    _com_interfaces_ = [IPolicyConfig]


# ---------------------------------------------------------------------------
# Notification callback
# ---------------------------------------------------------------------------
class _DeviceNotificationCallback(comtypes.COMObject):
    """COM callback that fires when devices change."""

    _com_interfaces_ = [IMMNotificationClient]

    def __init__(self, on_change: Callable[[], None]) -> None:
        super().__init__()
        self._on_change = on_change

    def OnDeviceStateChanged(self, pwstrDeviceId: str, dwNewState: int) -> HRESULT:
        self._on_change()
        return 0  # S_OK

    def OnDeviceAdded(self, pwstrDeviceId: str) -> HRESULT:
        self._on_change()
        return 0

    def OnDeviceRemoved(self, pwstrDeviceId: str) -> HRESULT:
        self._on_change()
        return 0

    def OnDefaultDeviceChanged(
        self, flow: int, role: int, pwstrDefaultDeviceId: str
    ) -> HRESULT:
        self._on_change()
        return 0

    def OnPropertyValueChanged(self, pwstrDeviceId: str, key: int) -> HRESULT:
        return 0  # Ignore property changes


# ---------------------------------------------------------------------------
# AudioManager
# ---------------------------------------------------------------------------
class AudioManager:
    """High-level interface for enumerating and switching audio devices."""

    def __init__(self) -> None:
        comtypes.CoInitializeEx(comtypes.COINIT_APARTMENTTHREADED)
        self._enumerator: IMMDeviceEnumerator = comtypes.CoCreateInstance(
            GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}"),
            IMMDeviceEnumerator,
        )
        self._callback: _DeviceNotificationCallback | None = None
        self._policy_config: IPolicyConfig | None = None

    def _get_policy_config(self) -> IPolicyConfig:
        """Lazily create and cache the IPolicyConfig COM object."""
        if self._policy_config is None:
            self._policy_config = comtypes.CoCreateInstance(
                CPolicyConfigClient._reg_clsid_,
                IPolicyConfig,
            )
        return self._policy_config

    def enumerate_devices(self) -> list[AudioDevice]:
        """Return all active audio input and output devices.

        Returns:
            A list of AudioDevice instances, with is_default set for
            the current default render and capture devices.
        """
        devices: list[AudioDevice] = []

        default_output_id = self._get_default_device_id(EDataFlow.eRender.value)
        default_input_id = self._get_default_device_id(EDataFlow.eCapture.value)

        for flow_val, flow_enum in [
            (EDataFlow.eRender.value, DeviceFlow.OUTPUT),
            (EDataFlow.eCapture.value, DeviceFlow.INPUT),
        ]:
            try:
                collection = self._enumerator.EnumAudioEndpoints(
                    flow_val, DEVICE_STATE.ACTIVE.value
                )
                count = collection.GetCount()
                for i in range(count):
                    dev: IMMDevice = collection.Item(i)
                    dev_id = dev.GetId()
                    props = dev.OpenPropertyStore(0)  # STGM_READ
                    try:
                        # PKEY_Device_FriendlyName = {a45c254e-df1c-4efd-8020-67d146a850e0}, 14
                        from comtypes import GUID as G
                        from pycaw.pycaw import PROPERTYKEY
                        pk = PROPERTYKEY()
                        pk.fmtid = G("{a45c254e-df1c-4efd-8020-67d146a850e0}")
                        pk.pid = 14
                        pv = props.GetValue(pk)
                        name = pv.GetValue() or f"Device {i}"
                    except Exception:
                        name = f"Device {i}"

                    default_id = (
                        default_output_id
                        if flow_enum == DeviceFlow.OUTPUT
                        else default_input_id
                    )
                    devices.append(
                        AudioDevice(
                            id=dev_id,
                            name=name,
                            flow=flow_enum,
                            is_default=(dev_id == default_id),
                        )
                    )
            except Exception as exc:
                log.warning("Failed to enumerate %s devices: %s", flow_enum.value, exc)

        return devices

    def _get_default_device_id(self, flow: int) -> str | None:
        """Return the device ID of the current default endpoint."""
        try:
            dev = self._enumerator.GetDefaultAudioEndpoint(flow, ERole.eMultimedia.value)
            return dev.GetId()
        except Exception:
            return None

    def set_default_device(self, device_id: str) -> bool:
        """Set the given device as the default for all roles.

        Args:
            device_id: The endpoint device ID string.

        Returns:
            True on success, False on failure.
        """
        try:
            policy = self._get_policy_config()
            for role in (ERole.eConsole.value, ERole.eMultimedia.value, ERole.eCommunications.value):
                policy.SetDefaultEndpoint(device_id, role)
            return True
        except Exception as exc:
            log.error("Failed to set default device %s: %s", device_id, exc)
            return self._fallback_set_default(device_id)

    def _fallback_set_default(self, device_id: str) -> bool:
        """Fallback: use AudioDeviceCmdlets via PowerShell."""
        import subprocess
        try:
            result = subprocess.run(
                [
                    "powershell",
                    "-Command",
                    f'Set-AudioDevice -ID "{device_id}"',
                ],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                timeout=5,
            )
            return result.returncode == 0
        except Exception as exc:
            log.error("Fallback device switch failed: %s", exc)
            return False

    def register_change_callback(self, on_change: Callable[[], None]) -> None:
        """Register a callback that fires when any device state changes.

        Args:
            on_change: A callable invoked (from a COM thread) on device changes.
        """
        if self._callback is not None:
            self.unregister_change_callback()
        self._callback = _DeviceNotificationCallback(on_change)
        self._enumerator.RegisterEndpointNotificationCallback(self._callback)

    def unregister_change_callback(self) -> None:
        """Remove the device change notification callback."""
        if self._callback is not None:
            try:
                self._enumerator.UnregisterEndpointNotificationCallback(self._callback)
            except Exception:
                pass
            self._callback = None

    def get_output_devices(self) -> list[AudioDevice]:
        """Convenience: return only output (render) devices."""
        return [d for d in self.enumerate_devices() if d.flow == DeviceFlow.OUTPUT]

    def get_input_devices(self) -> list[AudioDevice]:
        """Convenience: return only input (capture) devices."""
        return [d for d in self.enumerate_devices() if d.flow == DeviceFlow.INPUT]

    def get_default_output(self) -> AudioDevice | None:
        """Return the current default output device, or None."""
        for d in self.enumerate_devices():
            if d.flow == DeviceFlow.OUTPUT and d.is_default:
                return d
        return None

    def get_default_input(self) -> AudioDevice | None:
        """Return the current default input device, or None."""
        for d in self.enumerate_devices():
            if d.flow == DeviceFlow.INPUT and d.is_default:
                return d
        return None

    def get_default_volume(self, flow: DeviceFlow = DeviceFlow.OUTPUT) -> float | None:
        """Get master volume scalar (0.0-1.0) of the default device."""
        try:
            edata = EDataFlow.eRender.value if flow == DeviceFlow.OUTPUT else EDataFlow.eCapture.value
            dev = self._enumerator.GetDefaultAudioEndpoint(edata, ERole.eMultimedia.value)
            interface = dev.Activate(IAudioEndpointVolume._iid_, _CLSCTX_ALL, None)
            vol = cast(interface, POINTER(IAudioEndpointVolume))
            return vol.GetMasterVolumeLevelScalar()
        except Exception as exc:
            log.warning("Failed to get volume: %s", exc)
            return None

    def set_default_volume(self, level: float, flow: DeviceFlow = DeviceFlow.OUTPUT) -> bool:
        """Set master volume scalar (0.0-1.0) of the default device."""
        try:
            level = max(0.0, min(1.0, level))
            edata = EDataFlow.eRender.value if flow == DeviceFlow.OUTPUT else EDataFlow.eCapture.value
            dev = self._enumerator.GetDefaultAudioEndpoint(edata, ERole.eMultimedia.value)
            interface = dev.Activate(IAudioEndpointVolume._iid_, _CLSCTX_ALL, None)
            vol = cast(interface, POINTER(IAudioEndpointVolume))
            vol.SetMasterVolumeLevelScalar(level, None)
            return True
        except Exception as exc:
            log.error("Failed to set volume: %s", exc)
            return False
