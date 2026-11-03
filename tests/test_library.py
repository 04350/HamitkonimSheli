# -*- coding: utf-8 -*-
import sys
import unittest
import os
import tempfile
import shutil
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import ספריית_המתכונים as app
from PIL import Image, PngImagePlugin
from openpyxl import Workbook
from pptx import Presentation

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

    def test_זיהוי_סוגי_קבצים_נוספים(self):
        self.assertIn(".doc", app.OFFICE_TYPES)
        self.assertIn(".xls", app.OFFICE_TYPES)
        self.assertIn(".ppt", app.OFFICE_TYPES)
        self.assertIn(".docm", app.OFFICE_TYPES)
        self.assertIn(".ods", app.OFFICE_TYPES)
        self.assertEqual(app.file_kind(".xls"), "גיליון")
        self.assertEqual(app.file_kind(".ppt"), "מצגת")
        self.assertEqual(app.file_kind(".mobi"), "ספר דיגיטלי")
        self.assertIn(".epub", app.DOCUMENT_TYPES)

    def test_ספר_דיגיטלי_epub_נפתח_ונקרא(self):
        if app.fitz is None:
            self.skipTest("נדרשת ספריית PyMuPDF")
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / "בדיקה.epub"
            container_xml = (
                '<?xml version="1.0"?>'
                '<container xmlns="urn:oasis:names:tc:opendocument:xmlns:container" version="1.0">'
                '<rootfiles><rootfile full-path="OEBPS/content.opf" '
                'media-type="application/oebps-package+xml"/></rootfiles></container>'
            )
            package_xml = (
                '<?xml version="1.0" encoding="utf-8"?>'
                '<package xmlns="http://www.idpf.org/2007/opf" version="2.0" unique-identifier="bookid">'
                '<metadata xmlns:dc="http://purl.org/dc/elements/1.1/">'
                '<dc:title>ספר בדיקה</dc:title></metadata>'
                '<manifest><item id="chapter" href="chapter.xhtml" '
                'media-type="application/xhtml+xml"/></manifest>'
                '<spine><itemref idref="chapter"/></spine></package>'
            )
            chapter = (
                '<?xml version="1.0" encoding="utf-8"?>'
                '<html xmlns="http://www.w3.org/1999/xhtml"><head><title>בדיקה</title></head>'
                '<body><h1>עוגת תפוחים</h1><p>זהו טקסט בדיקה מתוך ספר דיגיטלי.</p></body></html>'
            )
            with app.zipfile.ZipFile(source, "w") as archive:
                archive.writestr("mimetype", "application/epub+zip", compress_type=app.zipfile.ZIP_STORED)
                archive.writestr("META-INF/container.xml", container_xml)
                archive.writestr("OEBPS/content.opf", package_xml)
                archive.writestr("OEBPS/chapter.xhtml", chapter)
            with app.fitz.open(source) as ebook:
                self.assertGreater(ebook.page_count, 0)
            extracted = app.extract_text(source)
            self.assertIn("עוגת תפוחים", extracted)

    @unittest.skipUnless(app.find_libreoffice() is not None, "נדרשת התקנת LibreOffice (קיימת בחבילת Windows)")
    def test_המרת_word_excel_powerpoint_ל_pdf(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            word_file = folder / "בדיקת Word.docx"
            word = app.Document()
            word.add_paragraph("WORD PREVIEW CHECK")
            word.save(word_file)

            excel_file = folder / "בדיקת Excel.xlsx"
            workbook = Workbook()
            workbook.active["A1"] = "EXCEL PREVIEW CHECK"
            workbook.save(excel_file)

            powerpoint_file = folder / "בדיקת PowerPoint.pptx"
            presentation = Presentation()
            slide = presentation.slides.add_slide(presentation.slide_layouts[5])
            box = slide.shapes.add_textbox(100, 100, 500, 100)
            box.text = "POWERPOINT PREVIEW CHECK"
            presentation.save(powerpoint_file)

            for source in (word_file, excel_file, powerpoint_file):
                with self.subTest(file=source.suffix):
                    pdf_path, temp_dir, error = app.convert_office_to_pdf(source)
                    self.assertIsNone(error, error)
                    self.assertIsNotNone(pdf_path)
                    self.assertTrue(pdf_path.is_file())
                    with app.fitz.open(pdf_path) as pdf:
                        self.assertGreater(pdf.page_count, 0)
                    shutil.rmtree(temp_dir, ignore_errors=True)

    def test_קובץ_office_פגום_לא_מפיל_את_ההמרה(self):
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / "קובץ.docx"
            source.write_bytes(b"this is not a valid Office document")
            with patch.object(app, "find_libreoffice", return_value=None):
                pdf_path, temp_dir, error = app.convert_office_to_pdf(source)
            self.assertIsNone(pdf_path)
            self.assertIsNone(temp_dir)
            self.assertTrue(error)

    def test_שם_ללא_תוכן_נשאר_קריא(self):
        self.assertGreater(app.relevance("קובץ תמונה", "", "תמונה"), 0)

if __name__ == "__main__":
    unittest.main()
