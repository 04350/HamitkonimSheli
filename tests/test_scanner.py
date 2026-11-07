# -*- coding: utf-8 -*-
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from recipe_scanner import lines_from_tesseract_data, select_title_line


class RecipeScannerTests(unittest.TestCase):
    def test_title_prefers_largest_printed_line_not_first_line(self):
        lines = [
            {"text": "מתכון מיוחד", "top": 24, "height": 16, "confidence": 90},
            {"text": "עוגת שוקולד עשירה", "top": 120, "height": 42, "confidence": 92},
            {"text": "מצרכים", "top": 220, "height": 18, "confidence": 95},
            {"text": "2 כוסות קמח", "top": 260, "height": 17, "confidence": 90},
        ]
        self.assertEqual(select_title_line(lines, 1000), "עוגת שוקולד עשירה")

    def test_ignores_section_headers_and_page_numbers(self):
        lines = [
            {"text": "מצרכים", "top": 20, "height": 40, "confidence": 98},
            {"text": "123", "top": 40, "height": 60, "confidence": 99},
            {"text": "עוגיות חמאה", "top": 130, "height": 30, "confidence": 85},
        ]
        self.assertEqual(select_title_line(lines, 900), "עוגיות חמאה")

    def test_rebuilds_hebrew_lines_in_right_to_left_order(self):
        data = {
            "text": ["שוקולד", "עוגת"],
            "conf": ["90", "90"],
            "left": ["10", "90"],
            "top": ["20", "20"],
            "width": ["70", "50"],
            "height": ["22", "22"],
            "block_num": ["1", "1"],
            "par_num": ["1", "1"],
            "line_num": ["1", "1"],
        }
        result = lines_from_tesseract_data(data)
        self.assertEqual(result[0]["text"], "עוגת שוקולד")

    def test_pdf_title_uses_larger_font_and_not_first_line(self):
        from recipe_scanner import best_title_from_pdf_page
        class FakePage:
            rect = type("Rect", (), {"height": 900})()
            def get_text(self, kind):
                return {"blocks": [{
                    "lines": [
                        {"bbox": [20, 20, 300, 34], "spans": [{"text": "ספר מתכונים", "size": 8}]},
                        {"bbox": [20, 120, 450, 160], "spans": [{"text": "עוגת שוקולד עשירה", "size": 22}]},
                        {"bbox": [20, 200, 300, 216], "spans": [{"text": "מצרכים", "size": 9}]},
                    ]
                }]}
        self.assertEqual(best_title_from_pdf_page(FakePage()), "עוגת שוקולד עשירה")


if __name__ == "__main__":
    unittest.main()
