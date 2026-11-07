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

    def test_screenshot_is_process_specific_and_checks_demo_data(self):
        capture = (ROOT / "scripts" / "capture_app_screenshot.py").read_text(encoding="utf-8")
        workflow = (ROOT / ".github" / "workflows" / "build-windows.yml").read_text(encoding="utf-8")
        self.assertIn("--process-id", capture)
        self.assertIn("--expected-items", capture)
        self.assertIn('Environment["LOCALAPPDATA"]', workflow)
        self.assertNotIn("continue-on-error: true", workflow)


if __name__ == "__main__":
    unittest.main()
