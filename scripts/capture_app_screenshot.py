# -*- coding: utf-8 -*-
"""Capture a real screenshot of the running Windows desktop application."""
from __future__ import annotations

import argparse
import ctypes
import time
from ctypes import wintypes
from pathlib import Path

from PIL import ImageGrab, ImageStat


def process_family_ids(root_pid: int) -> set[int]:
    """Return the launched process and its descendants (PyInstaller --onefile may spawn a child)."""
    from ctypes import wintypes

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD),
            ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.c_size_t),
            ("th32ModuleID", wintypes.DWORD),
            ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", wintypes.LONG),
            ("dwFlags", wintypes.DWORD),
            ("szExeFile", wintypes.WCHAR * 260),
        ]

    kernel32 = ctypes.windll.kernel32
    kernel32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
    kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    kernel32.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    kernel32.Process32FirstW.restype = wintypes.BOOL
    kernel32.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    kernel32.Process32NextW.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL

    snapshot = kernel32.CreateToolhelp32Snapshot(0x00000002, 0)  # TH32CS_SNAPPROCESS
    snapshot_value = snapshot if isinstance(snapshot, int) else snapshot.value
    invalid_handle = ctypes.c_void_p(-1).value
    if snapshot_value == invalid_handle or not snapshot_value:
        raise OSError("Could not enumerate Windows processes.")

    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        parents: dict[int, int] = {}
        if not kernel32.Process32FirstW(snapshot, ctypes.byref(entry)):
            raise OSError("Could not read the Windows process list.")
        while True:
            parents[int(entry.th32ProcessID)] = int(entry.th32ParentProcessID)
            if not kernel32.Process32NextW(snapshot, ctypes.byref(entry)):
                break

        family = {int(root_pid)}
        changed = True
        while changed:
            changed = False
            for pid, parent_pid in parents.items():
                if parent_pid in family and pid not in family:
                    family.add(pid)
                    changed = True
        return family
    finally:
        kernel32.CloseHandle(snapshot)


def find_app_window(title_fragment: str, expected_pid: int | None = None):
    user32 = ctypes.windll.user32
    matches = []
    allowed_pids = process_family_ids(expected_pid) if expected_pid is not None else None
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
            if allowed_pids is None or int(process_id.value) in allowed_pids:
                matches.append((hwnd, title, int(process_id.value)))
        return True

    callback_ref = enum_proc_type(callback)
    user32.EnumWindows(callback_ref, 0)
    if not matches:
        return (None, "")
    # Prefer the window owned by the original process; otherwise use its child.
    matches.sort(key=lambda item: item[2] != expected_pid if expected_pid is not None else False)
    hwnd, title, _pid = matches[0]
    return hwnd, title


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
