# -*- coding: utf-8 -*-
import sys
import unittest
from pathlib import Path
from unittest.mock import Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from PIL import Image
from recipe_scanner import (
    MAX_OCR_EDGE, MAX_OCR_PIXELS, best_title_from_pdf_page,
    lines_from_tesseract_data, preprocess_image, recognize_image, select_title_line,
)


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
            "text": ["שוקולד", "עוגת"], "conf": ["90", "90"],
            "left": ["10", "90"], "top": ["20", "20"], "width": ["70", "50"],
            "height": ["22", "22"], "block_num": ["1", "1"],
            "par_num": ["1", "1"], "line_num": ["1", "1"],
        }
        result = lines_from_tesseract_data(data)
        self.assertEqual(result[0]["text"], "עוגת שוקולד")

    def test_pdf_title_uses_larger_font_and_not_first_line(self):
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

    def test_preprocess_downsamples_large_images_before_ocr(self):
        image = Image.new("RGB", (6000, 4000), "white")
        enhanced, threshold = preprocess_image(image)
        try:
            width, height = enhanced.size
            self.assertLessEqual(max(width, height), MAX_OCR_EDGE)
            self.assertLessEqual(width * height, MAX_OCR_PIXELS)
            self.assertEqual(enhanced.mode, "L")
        finally:
            enhanced.close()
            threshold.close()
            image.close()

    def test_good_first_pass_runs_only_one_ocr_variant(self):
        class FakeOutput:
            DICT = "dict"
        fake = Mock()
        fake.Output = FakeOutput
        fake.image_to_data.return_value = {
            "text": ["APPLE", "RECIPE"], "conf": ["96", "94"],
            "left": ["10", "140"], "top": ["20", "20"],
            "width": ["110", "120"], "height": ["42", "42"],
            "block_num": ["1", "1"], "par_num": ["1", "1"], "line_num": ["1", "1"],
        }
        image = Image.new("RGB", (900, 400), "white")
        try:
            result = recognize_image(image, fake)
        finally:
            image.close()
        self.assertIn("APPLE", result["text"])
        self.assertEqual(fake.image_to_data.call_count, 1)

    def test_low_confidence_first_pass_runs_one_retry(self):
        class FakeOutput:
            DICT = "dict"
        fake = Mock()
        fake.Output = FakeOutput
        fake.image_to_data.side_effect = [
            {"text": ["x"], "conf": ["5"], "left": ["1"], "top": ["1"], "width": ["1"], "height": ["1"],
             "block_num": ["1"], "par_num": ["1"], "line_num": ["1"]},
            {"text": ["עוגת", "תפוחים"], "conf": ["91", "92"], "left": ["10", "100"],
             "top": ["20", "20"], "width": ["80", "90"], "height": ["35", "35"],
             "block_num": ["1", "1"], "par_num": ["1", "1"], "line_num": ["1", "1"]},
        ]
        image = Image.new("RGB", (900, 400), "white")
        try:
            result = recognize_image(image, fake)
        finally:
            image.close()
        self.assertIn("עוגת", result["text"])
        self.assertEqual(fake.image_to_data.call_count, 2)


if __name__ == "__main__":
    unittest.main()
