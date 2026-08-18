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

from src.icons import match_icon_for_name  # noqa: E402
from src.ui import _bt_name_core, _bt_names_match  # noqa: E402


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
