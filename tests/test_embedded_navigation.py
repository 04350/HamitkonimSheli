# -*- coding: utf-8 -*-
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class EmbeddedNavigationTests(unittest.TestCase):
    def test_primary_screens_are_embedded_in_main_window(self):
        source = (ROOT / "src" / "ספריית_המתכונים.py").read_text(encoding="utf-8")
        self.assertNotIn("tk.Toplevel(", source)
        for expected in (
            "def _push_embedded_page(",
            "def _pop_embedded_page(",
            'window = self._push_embedded_page("סורק מתכונים חכם")',
            'dialog = self._push_embedded_page("הגדרות")',
            'dialog = self._push_embedded_page("סל המחזור")',
            'progress_window = self._push_embedded_page("ייבוא ספרייה")',
            'window = self._push_embedded_page(title)',
        ):
            self.assertIn(expected, source)

    def test_recipe_center_uses_the_same_embedded_stack(self):
        source = (ROOT / "src" / "recipe_center.py").read_text(encoding="utf-8")
        self.assertIn('window = app._push_embedded_page("מרכז המתכונים")', source)
        self.assertNotIn("tk.Toplevel(app.root)", source)

    def test_recipe_reader_can_switch_between_image_and_extracted_text(self):
        source = (ROOT / "src" / "ספריית_המתכונים.py").read_text(encoding="utf-8")
        self.assertIn('self.reader_preview_mode = tk.StringVar(', source)
        self.assertIn('command=lambda: self._set_reader_preview_mode("image")', source)
        self.assertIn('command=lambda: self._set_reader_preview_mode("text")', source)
        self.assertIn('source_text = str(row["content"] or "").strip()', source)

    def test_reader_panel_is_draggable_and_remembers_width(self):
        source = (ROOT / "src" / "ספריית_המתכונים.py").read_text(encoding="utf-8")
        self.assertIn("self.content_splitter = tk.PanedWindow(", source)
        self.assertIn('self.content_splitter.bind("<ButtonRelease-1>", self._save_reader_sash_position)', source)
        self.assertIn('"reader_pane_width"', source)

    def test_file_viewer_is_an_embedded_page_with_source_and_text_modes(self):
        source = (ROOT / "src" / "ספריית_המתכונים.py").read_text(encoding="utf-8")
        self.assertIn("self._embedded_header(", source)
        self.assertIn('window, title, "תצוגת המקור', source)
        self.assertIn('text="▧  תמונה / מסמך מקורי"', source)
        self.assertIn('text="☷  טקסט שחולץ"', source)


if __name__ == "__main__":
    unittest.main()
