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
from ctypes import POINTER, c_uint, cast
from ctypes.wintypes import LPCWSTR, BOOL

from pycaw.pycaw import (
    IAudioEndpointVolume,
    IMMDeviceEnumerator,
    IMMDevice,
    IMMNotificationClient,
    EDataFlow,
    ERole,
    DEVICE_STATE,
)

log = logging.getLogger(__name__)

_PKEY_FMTID = "{a45c254e-df1c-4efd-8020-67d146a850e0}"
_PKEY_DEVICE_FRIENDLY_NAME = 14
_PKEY_DEVICE_ENUMERATOR_NAME = 24

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
    is_bluetooth: bool = False
    is_connected: bool = True  # False only for "ghost" entries from config


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

    @staticmethod
    def _get_prop(props: object, pid: int) -> str:
        """Read one string property from an already-open property store."""
        from comtypes import GUID as G
        from pycaw.pycaw import PROPERTYKEY
        pk = PROPERTYKEY()
        pk.fmtid = G(_PKEY_FMTID)
        pk.pid = pid
        return props.GetValue(pk).GetValue() or ""

    def _read_device_props(self, dev: IMMDevice, fallback: str = "Unknown") -> tuple[str, bool]:
        """Return (friendly_name, is_bluetooth) for one endpoint.

        Shared by enumerate_devices() and get_default_device() so the
        Bluetooth detection rules stay in one place.
        """
        props = dev.OpenPropertyStore(0)  # STGM_READ

        try:
            name = self._get_prop(props, _PKEY_DEVICE_FRIENDLY_NAME) or fallback
        except Exception:
            name = fallback

        # Primary signal: PKEY_Device_EnumeratorName. Some stacks report
        # BTHENUM, others BTHHFENUM, hence the prefix test.
        enumerator_name = ""
        is_bt = False
        try:
            enumerator_name = self._get_prop(props, _PKEY_DEVICE_ENUMERATOR_NAME).upper()
            is_bt = enumerator_name.startswith("BTH")
        except Exception as exc:
            log.debug("Enumerator detection failed for '%s': %s", name, exc)

        # Fallback: cross-reference against paired BT device names. Intel/Dell
        # audio controllers proxy BT audio through their own driver, reporting
        # enumerator='INTELAUDIO' instead of a BTH* prefix. The paired-name
        # list is cached for 30s, so this stays cheap under polling.
        if not is_bt:
            try:
                from .bluetooth import get_paired_device_names
                name_lower = name.lower()
                for bt_name in get_paired_device_names():
                    bt_lower = bt_name.lower()
                    if bt_lower in name_lower or name_lower in bt_lower:
                        is_bt = True
                        log.info(
                            "Device '%s' matched paired BT device '%s' "
                            "(name cross-ref, enumerator='%s')",
                            name, bt_name, enumerator_name,
                        )
                        break
            except Exception as exc:
                log.debug("BT name cross-ref failed for '%s': %s", name, exc)

        log.debug("Device '%s' enumerator='%s' is_bt=%s", name, enumerator_name, is_bt)
        return name, is_bt

    def get_default_device(self, flow: DeviceFlow = DeviceFlow.OUTPUT) -> AudioDevice | None:
        """Return the current default endpoint for *flow*, or None.

        Asks Windows for the default endpoint directly instead of enumerating
        every device and filtering. The widget polls this every 2s, where the
        full enumeration was doing several times the necessary COM work.
        """
        try:
            edata = (
                EDataFlow.eRender.value
                if flow == DeviceFlow.OUTPUT
                else EDataFlow.eCapture.value
            )
            dev = self._enumerator.GetDefaultAudioEndpoint(
                edata, ERole.eMultimedia.value
            )
            name, is_bt = self._read_device_props(dev)
            return AudioDevice(
                id=dev.GetId(), name=name, flow=flow,
                is_default=True, is_bluetooth=is_bt,
            )
        except Exception as exc:
            log.debug("No default %s device: %s", flow.value, exc)
            return None

    def enumerate_devices(self) -> list[AudioDevice]:
        """Return all active audio input and output devices.

        Returns:
            A list of AudioDevice instances, with is_default set for
            the current default render and capture devices.
        """
        log.debug("enumerate_devices() called")
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
                    name, is_bt = self._read_device_props(dev, fallback=f"Device {i}")

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
                            is_bluetooth=is_bt,
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
            # There is deliberately no fallback here. The previous one shelled
            # out to Set-AudioDevice from the third-party AudioDeviceCmdlets
            # module, which is not a dependency and is not installed by
            # default, so it spawned a PowerShell process only to fail.
            log.error("Failed to set default device %s: %s", device_id, exc)
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
        return self.get_default_device(DeviceFlow.OUTPUT)

    def get_default_input(self) -> AudioDevice | None:
        """Return the current default input device, or None."""
        return self.get_default_device(DeviceFlow.INPUT)

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
