"""Tests for Bluetooth endpoint name matching and icon keyword matching.

Dependency-free: run with `python tests/test_name_matching.py`.

These are the functions behind BT favourite reconciliation - when a device
reconnects with a new endpoint ID, the only thing tying the old favourite to
the new endpoint is its name - so getting them wrong silently orphans a
favourite and its icon override.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.audio_manager import _filter_path_from_device_id  # noqa: E402
from src.icons import match_icon_for_name  # noqa: E402
from src.ui import _bt_name_core, _bt_names_match  # noqa: E402

# Real topology device IDs captured from a Buds Pro 2 and a Realtek card
_A2DP_ID = (
    "{2}.\\\\?\\bthenum#{0000110b-0000-1000-8000-00805f9b34fb}_vid&0001e000_pid&0500"
    "#7&c600d85&0&fc915d65a039_c00000000#{6994ad04-93ef-11d0-a3cc-00a0c9223196}\\src"
)
_HFP_ID = (
    "{2}.\\\\?\\bthhfenum#bthhfpaudio#8&21fac373&0&97"
    "#{6994ad04-93ef-11d0-a3cc-00a0c9223196}\\wave"
)
_HDAUDIO_ID = (
    "{2}.\\\\?\\hdaudio#func_01&ven_10ec&dev_0897&subsys_1462d812&rev_1001"
    "#5&1e3c8a3f&0&0001#{6994ad04-93ef-11d0-a3cc-00a0c9223196}\\eheadphonetopo"
)


def test_filter_path_strips_prefix_for_a2dp() -> None:
    assert _filter_path_from_device_id(_A2DP_ID) == _A2DP_ID[len("{2}."):]


def test_filter_path_accepts_hands_free() -> None:
    assert _filter_path_from_device_id(_HFP_ID) == _HFP_ID[len("{2}."):]


def test_filter_path_rejects_non_bluetooth() -> None:
    # A Bluetooth request must never be sent to an ordinary sound card
    assert _filter_path_from_device_id(_HDAUDIO_ID) is None
    assert _filter_path_from_device_id("") is None


def test_core_extracts_parenthesised_name() -> None:
    assert _bt_name_core("Headphones (Buds Pro 2)") == "buds pro 2"
    assert _bt_name_core("Earphones (Buds Pro 2)") == "buds pro 2"


def test_core_uses_last_parentheses() -> None:
    # Windows nests these: "Audio out (headphones!) (High Definition Audio Device)"
    assert _bt_name_core(
        "Audio out (headphones!) (High Definition Audio Device)"
    ) == "high definition audio device"


def test_core_falls_back_to_whole_name() -> None:
    assert _bt_name_core("Speakers") == "speakers"
    assert _bt_name_core("  Speakers  ") == "speakers"


def test_core_handles_unbalanced_parentheses() -> None:
    assert _bt_name_core("Headphones Buds Pro 2)") == "headphones buds pro 2)"


def test_same_device_different_endpoint_roles_match() -> None:
    # The case reconciliation exists for: one BT device, two endpoint names
    assert _bt_names_match("Headphones (Buds Pro 2)", "Earphones (Buds Pro 2)")


def test_different_devices_do_not_match() -> None:
    assert not _bt_names_match("Headphones (Buds Pro 2)", "Headphones (WH-1000XM4)")


def test_empty_names_do_not_match() -> None:
    # Guards against every unnamed device matching every other one
    assert not _bt_names_match("", "")
    assert not _bt_names_match("()", "()")


def test_icon_keyword_matching() -> None:
    assert match_icon_for_name("Headphones (Buds Pro 2)") == "headphones"
    assert match_icon_for_name("Realtek High Definition Audio") == "speaker"
    assert match_icon_for_name("Samsung TV") == "tv"
    assert match_icon_for_name("USB Audio Device") == "usb"


def test_icon_falls_back_for_unknown_names() -> None:
    assert match_icon_for_name("Zzzz Unknown Thing") == "audio"


def test_icon_matching_is_case_insensitive() -> None:
    assert match_icon_for_name("HEADPHONES") == "headphones"
    assert match_icon_for_name("headphones") == "headphones"


if __name__ == "__main__":
    failures = 0
    for name, fn in sorted(globals().items()):
        if not name.startswith("test_") or not callable(fn):
            continue
        try:
            fn()
            print(f"  PASS  {name}")
        except AssertionError as exc:
            failures += 1
            print(f"  FAIL  {name}  {exc}")
    print(f"\n{'FAILED' if failures else 'All passed'} ({failures} failure(s))")
    sys.exit(1 if failures else 0)
