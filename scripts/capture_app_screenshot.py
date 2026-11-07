# -*- coding: utf-8 -*-
"""Capture a real screenshot of the running Windows desktop application."""
from __future__ import annotations

import argparse
import ctypes
import time
from ctypes import wintypes
from pathlib import Path

from PIL import ImageGrab, ImageStat


def find_app_window(title_fragment: str, expected_pid: int | None = None):

    user32 = ctypes.windll.user32
    matches = []
    enum_proc_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def callback(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return True
        buffer = ctypes.create_unicode_buffer(length + 2)
        user32.GetWindowTextW(hwnd, buffer, len(buffer))
        title = buffer.value
        if title_fragment in title:
            process_id = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
            if expected_pid is None or process_id.value == expected_pid:
                matches.append((hwnd, title))
        return True

    callback_ref = enum_proc_type(callback)
    user32.EnumWindows(callback_ref, 0)
    return matches[0] if matches else (None, "")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True)
    parser.add_argument("--process-id", type=int, required=True, help="PID returned by Start-Process")
    parser.add_argument("--expected-items", type=int, default=0, help="Minimum number of demo records that must exist")
    args = parser.parse_args()

    if args.expected_items < 0:
        raise SystemExit("--expected-items cannot be negative")

    # Verify the isolated demo profile contains the expected records before
    # capturing anything, so a splash screen or empty library is not accepted.
    import sys
    project_root = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(project_root / "src"))
    import ספריית_המתכונים as app

    library = app.RecipeLibrary()
    try:
        actual_items = len(library.rows())
    finally:
        library.db.close()
    if actual_items < args.expected_items:
        raise SystemExit(
            f"Demo library is incomplete: found {actual_items}; expected at least {args.expected_items}"
        )

    user32 = ctypes.windll.user32
    hwnd, title = find_app_window("ספריית המתכונים שלי", args.process_id)
    if not hwnd:
        raise SystemExit("The running recipe-library window was not found.")
    user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
    user32.SetForegroundWindow(hwnd)
    time.sleep(2.0)

    rect = wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        raise SystemExit("Could not read the application window rectangle.")
    left, top, right, bottom = rect.left, rect.top, rect.right, rect.bottom
    if right - left < 900 or bottom - top < 600:
        raise SystemExit(f"Unexpected app window size: {right-left}x{bottom-top}")

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    image = ImageGrab.grab(
        bbox=(left, top, right, bottom),
        all_screens=True,
        include_layered_windows=True,
    )
    if image.width < 900 or image.height < 600:
        raise SystemExit(f"Captured image is too small: {image.size}")
    deviations = ImageStat.Stat(image.convert("RGB")).stddev
    if max(deviations) < 8:
        raise SystemExit("Captured application window appears blank or uniform.")
    image.save(output, format="PNG")
    print(
        f"Actual application window captured successfully: "
        f"{image.width}x{image.height}; demo_items={actual_items}; output={output}"
    )


if __name__ == "__main__":
    main()
