"""Tests for fullscreen detection geometry.

Dependency-free: run with `python tests/test_fullscreen.py` (pytest also
picks it up if you have it). Kept separate from the Win32 calls so the
decision logic can be checked without a real fullscreen window, which
Windows will not let a background process create reliably.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.ui import covers_monitor  # noqa: E402

# The 5120x1440 ultrawide this was developed against. Win32 RECTs, so right
# and bottom are exclusive.
MONITOR = (0, 0, 5120, 1440)
# Taskbar takes the bottom 48px, so a maximised window stops at 1392
MAXIMISED = (0, 0, 5120, 1392)


def test_exact_fullscreen() -> None:
    assert covers_monitor(MONITOR, MONITOR)


def test_maximised_is_not_fullscreen() -> None:
    # The important one: a maximised window covers rcWork but leaves the
    # taskbar visible, and must not trigger the yield.
    assert not covers_monitor(MAXIMISED, MONITOR)


def test_windowed_is_not_fullscreen() -> None:
    # A real Chrome window rect captured on the dev machine
    assert not covers_monitor((1524, 12, 3594, 1383), MONITOR)


def test_within_tolerance() -> None:
    # Apps that miss the edge by a pixel still count
    assert covers_monitor((0, 0, 5120, 1439), MONITOR)
    assert covers_monitor((1, 1, 5119, 1439), MONITOR)


def test_outside_tolerance() -> None:
    assert not covers_monitor((0, 0, 5120, 1437), MONITOR)
    assert not covers_monitor((0, 0, 5117, 1440), MONITOR)


def test_oversized_window_counts() -> None:
    # Some games size themselves slightly larger than the display
    assert covers_monitor((-8, -8, 5128, 1448), MONITOR)


def test_second_monitor() -> None:
    # A monitor to the right of the primary, at a negative-free offset
    second = (5120, 0, 7040, 1080)
    assert covers_monitor(second, second)
    assert not covers_monitor(MONITOR, second)


def test_empty_window_is_not_fullscreen() -> None:
    assert not covers_monitor((0, 0, 0, 0), MONITOR)


if __name__ == "__main__":
    failures = 0
    for name, fn in sorted(globals().items()):
        if not name.startswith("test_") or not callable(fn):
            continue
        try:
            fn()
            print(f"  PASS  {name}")
        except AssertionError:
            failures += 1
            print(f"  FAIL  {name}")
    print(f"\n{'FAILED' if failures else 'All passed'} ({failures} failure(s))")
    sys.exit(1 if failures else 0)
