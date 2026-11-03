# -*- coding: utf-8 -*-
import sys
import unittest
import os
import tempfile
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import ספריית_המתכונים as app
from PIL import Image, PngImagePlugin

class בדיקותספרייה(unittest.TestCase):
    def test_ניקוי_חיפוש_עברית_וניקוד(self):
        self.assertEqual(app.normalize("עוגת שׁוֹקוֹלָד!"), app.normalize("עוגת שוקולד"))

    def test_שם_קובץ_בטוח(self):
        self.assertEqual(app.safe_title('עוגה: "טעימה"?'), "עוגה טעימה")

    def test_זיהוי_כותרת_מתוך_טקסט(self):
        self.assertEqual(app.choose_title(Path("קובץ.txt"), "עוגת תפוחים\nמערבבים את החומרים"), "עוגת תפוחים")

    def test_חיפוש_נותן_עדיפות_לשם(self):
        name_score = app.relevance("עוגת שוקולד", "מתכון פשוט", "עוגת שוקולד")
        body_score = app.relevance("מתכון אחר", "עוגת שוקולד טעימה", "עוגת שוקולד")
        self.assertGreater(name_score, body_score)

    def test_ייבוא_תמונה_עם_כותרת_ומניעת_כפילות(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = temp
                source = Path(temp) / "תמונה.png"
                metadata = PngImagePlugin.PngInfo()
                metadata.add_text("Title", "עוגת גבינה")
                Image.new("RGB", (20, 20), "white").save(source, pnginfo=metadata)
                library = app.RecipeLibrary()
                added, title = library.import_one(source)
                self.assertTrue(added)
                self.assertEqual(title, "עוגת גבינה")
                self.assertEqual(library.rows()[0]["extension"], ".png")
                duplicate, _ = library.import_one(source)
                self.assertFalse(duplicate)
                library.db.close()
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous

    def test_סוגי_קבצים_מוצגים_בעברית(self):
        self.assertEqual(app.file_kind(".png"), "תמונה")
        self.assertEqual(app.file_kind(".pdf"), "מסמך")
        self.assertEqual(app.file_kind(".xlsx"), "גיליון")

    def test_זיהוי_טקסט_מתמונה_ללא_מטאדאטה(self):
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / "סריקה.png"
            Image.new("RGB", (120, 80), "white").save(source)
            with patch.object(app, "ocr_image", return_value="עוגת תפוחים") as ocr:
                self.assertEqual(app.extract_text(source), "עוגת תפוחים")
                ocr.assert_called_once()

    @unittest.skipIf(app.fitz is None or app.Image is None, "נדרשות ספריות PDF ותמונה")
    def test_זיהוי_טקסט_במסמך_pdf_סרוק(self):
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / "סריקה.pdf"
            document = app.fitz.open()
            document.new_page()
            document.save(str(source))
            document.close()
            with patch.object(app, "ocr_image", return_value="מתכון סרוק") as ocr:
                self.assertEqual(app.extract_text(source), "מתכון סרוק")
                ocr.assert_called_once()

    def test_שם_ללא_תוכן_נשאר_קריא(self):
        self.assertGreater(app.relevance("קובץ תמונה", "", "תמונה"), 0)

if __name__ == "__main__":
    unittest.main()
