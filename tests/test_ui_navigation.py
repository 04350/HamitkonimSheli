# -*- coding: utf-8 -*-
"""Static regression checks for the single-window recipe-library UX."""
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class InAppNavigationTests(unittest.TestCase):
    def test_primary_screens_are_embedded_and_sidebar_is_persistent(self):
        source = (ROOT / "src" / "ספריית_המתכונים.py").read_text(encoding="utf-8")
        center = (ROOT / "src" / "recipe_center.py").read_text(encoding="utf-8")
        self.assertIn("self.workspace = workspace", source)
        self.assertIn('host = getattr(self, "workspace", self.shell)', source)
        self.assertIn('app._push_embedded_page("מרכז המתכונים")', center)
        self.assertNotIn("tk.Toplevel(", source)
        self.assertNotIn("tk.Toplevel(", center)

    def test_reader_can_switch_preview_and_resize_the_divider(self):
        source = (ROOT / "src" / "ספריית_המתכונים.py").read_text(encoding="utf-8")
        self.assertIn("תמונה מקורית", source)
        self.assertIn("טקסט שחולץ", source)
        self.assertIn("sashwidth=13", source)
        self.assertIn("def _reset_reader_sash_position", source)
        self.assertIn("reader_pane_width", source)
        self.assertIn("source.thumbnail((304, 430)", source)

    def test_global_scrolling_covers_sidebar_workspace_and_embedded_pages(self):
        source = (ROOT / "src" / "ספריית_המתכונים.py").read_text(encoding="utf-8")
        for expected in (
            "class VerticalScrollFrame(tk.Frame):",
            "self.sidebar_scroll_frame = VerticalScrollFrame(",
            "self.workspace_scroll_frame = VerticalScrollFrame(",
            "page_scroll_frame = VerticalScrollFrame(",
            'for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):',
            'self.root.bind_all(sequence, self._route_mousewheel, add="+")',
            "def _route_mousewheel(self, event):",
            'widget = getattr(widget, "master", None)',
        ):
            self.assertIn(expected, source)

        self.assertNotIn('self.gallery_canvas.bind("<MouseWheel>"', source)
        self.assertIn('APP_VERSION = "2.0.0"', source)

    def test_screenshot_is_process_specific_and_checks_demo_data(self):
        capture = (ROOT / "scripts" / "capture_app_screenshot.py").read_text(encoding="utf-8")
        workflow = (ROOT / ".github" / "workflows" / "build-windows.yml").read_text(encoding="utf-8")
        self.assertIn("--process-id", capture)
        self.assertIn("--expected-items", capture)
        self.assertIn("args.process_id", capture)
        self.assertIn("actual_items < args.expected_items", capture)
        self.assertIn("def process_family_ids", capture)
        self.assertIn("TH32CS_SNAPPROCESS", capture)
        self.assertIn("CreateToolhelp32Snapshot.restype = wintypes.HANDLE", capture)
        self.assertIn('return (None, "")', capture)
        self.assertIn("process_id.value", capture)
        self.assertIn("capture-screenshot:", workflow)
        self.assertIn("needs: build-windows", workflow)
        self.assertIn("הרצה שנייה – צילום מסך לאחר בדיקות תקינות", workflow)
        self.assertIn('$env:LOCALAPPDATA = Join-Path', workflow)
        self.assertIn("--process-id $process.Id --expected-items 6", workflow)
        self.assertNotIn("continue-on-error: true", workflow)
        build_job, screenshot_job = workflow.split("  capture-screenshot:", 1)
        self.assertNotIn("צילום מסך אמיתי של הממשק מתוך Windows", build_job)
        self.assertIn("צילום מסך אמיתי של הממשק מתוך Windows", screenshot_job)

        demo_setup = (ROOT / "scripts" / "create_ui_demo_data.py").read_text(encoding="utf-8")
        self.assertIn('sys.stdout.reconfigure(encoding="utf-8", errors="replace")', demo_setup)


if __name__ == "__main__":
    unittest.main()
