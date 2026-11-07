# -*- coding: utf-8 -*-
"""Capture a real screenshot of the running Windows desktop application."""
from __future__ import annotations

import argparse
import ctypes
import time
from ctypes import wintypes
from pathlib import Path

from PIL import ImageGrab


def find_app_window(title_fragment: str):
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
            matches.append((hwnd, title))
        return True

    callback_ref = enum_proc_type(callback)
    user32.EnumWindows(callback_ref, 0)
    return matches[0] if matches else (None, "")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    user32 = ctypes.windll.user32
    hwnd, title = find_app_window("ספריית המתכונים שלי")
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
    image.save(output, format="PNG")
    print(f"Actual application window captured: {title} -> {output} ({image.width}x{image.height})")


if __name__ == "__main__":
    main()
