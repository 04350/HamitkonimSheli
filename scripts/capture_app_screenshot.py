# -*- coding: utf-8 -*-
"""Capture a real screenshot of the running Windows desktop application."""
from __future__ import annotations

import argparse
import ctypes
import os
import sqlite3
import time
from ctypes import wintypes
from pathlib import Path

from PIL import ImageGrab


def find_app_window(title_fragment: str, expected_process_id: int | None = None):
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
            if expected_process_id is not None:
                process_id = wintypes.DWORD()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
                if process_id.value != expected_process_id:
                    return True
            matches.append((hwnd, title))
        return True

    callback_ref = enum_proc_type(callback)
    user32.EnumWindows(callback_ref, 0)
    return matches[0] if matches else (None, "")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True)
    parser.add_argument("--process-id", type=int, default=None)
    parser.add_argument("--expected-items", type=int, default=0)
    args = parser.parse_args()

    if args.expected_items:
        profile = os.environ.get("LOCALAPPDATA")
        if not profile:
            raise SystemExit("LOCALAPPDATA is missing in screenshot process.")
        db_path = Path(profile) / "ספריית המתכונים שלי" / "הספרייה.sqlite3"
        if not db_path.is_file():
            raise SystemExit("Demo database was not created in the expected profile.")
        with sqlite3.connect(db_path) as connection:
            count = connection.execute("SELECT COUNT(*) FROM items WHERE is_deleted=0").fetchone()[0]
        if count < args.expected_items:
            raise SystemExit(f"Expected at least {args.expected_items} demo items; found {count}.")

    user32 = ctypes.windll.user32
    hwnd = None
    title = ""
    deadline = time.time() + 25
    while time.time() < deadline and hwnd is None:
        hwnd, title = find_app_window("ספריית המתכונים שלי", args.process_id)
        if hwnd is None:
            time.sleep(0.25)
    if not hwnd:
        raise SystemExit("The recipe-library window for the requested process was not found.")
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
    print(f"Actual application window captured (PID {args.process_id or 'any'}) -> {output} ({image.width}x{image.height})")


if __name__ == "__main__":
    main()
