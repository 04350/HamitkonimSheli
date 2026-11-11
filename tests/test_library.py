# -*- coding: utf-8 -*-
import sys
import unittest
import os
import tempfile
import time
import shutil
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import ספריית_המתכונים as app
from PIL import Image, PngImagePlugin
from openpyxl import Workbook
from pptx import Presentation

class בדיקותספרייה(unittest.TestCase):
    def test_libreoffice_חילוץ_עם_התקדמות_ועדכון_גרסה(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = temp
                root = Path(temp)
                first_archive = root / "office-v1.zip"
                second_archive = root / "office-v2.zip"

                def make_archive(target, version):
                    with app.zipfile.ZipFile(target, "w", compression=app.zipfile.ZIP_DEFLATED) as archive:
                        archive.writestr("program/soffice.exe", b"test-office-executable")
                        archive.writestr("program/soffice.com", b"test-office-console")
                        archive.writestr(".recipe-library-version", version)

                make_archive(first_archive, "26.2.6.1")
                progress = []
                executable = app._extract_portable_libreoffice(
                    first_archive,
                    progress_callback=lambda phase, current, total: progress.append((phase, current, total)) or True,
                )
                self.assertIsNotNone(executable)
                self.assertTrue(executable.is_file())
                self.assertEqual(progress[-1], ("extracting", 3, 3))

                installed = executable.parents[1]
                marker = installed / ".recipe-library-version"
                self.assertEqual(marker.read_text(encoding="utf-8"), "26.2.6.1")
                sentinel = installed / "keep-if-cache-hit.txt"
                sentinel.write_text("cached", encoding="utf-8")

                same_executable = app._extract_portable_libreoffice(first_archive)
                self.assertEqual(same_executable, executable)
                self.assertTrue(sentinel.exists(), "same-version cache hit should not re-extract files")

                make_archive(second_archive, "26.2.7.0")
                updated = app._extract_portable_libreoffice(second_archive)
                self.assertTrue(updated.is_file())
                self.assertFalse(sentinel.exists(), "a changed version should replace the old runtime")
                self.assertEqual(marker.read_text(encoding="utf-8"), "26.2.7.0")
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous

    def test_מטמון_קטגוריות_תיקיות_ואינדקסים_מתעדכן(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = str(Path(temp) / "profile")
                source = Path(temp) / "recipe.txt"
                source.write_text("עוגת תפוחים\nמצרכים:\n2 תפוחים\nאופן ההכנה:\nמערבבים", encoding="utf-8")
                library = app.RecipeLibrary()
                added, _ = library.import_one(source, relative_folder="קינוחים/עוגות")
                self.assertTrue(added)
                self.assertIn("קינוחים", library.all_source_folders())
                self.assertIn("קינוחים/עוגות", library.all_source_folders())
                self.assertIn("קינוחים/עוגות", library.all_categories())
                item_id = int(library.rows()[0]["id"])
                library.set_item_tags(item_id, ["חלבי", "שבת"])
                self.assertEqual(library.tags_by_item([item_id]), {item_id: ["חלבי", "שבת"]})

                indexes = {
                    row["name"] for row in library.db.execute("PRAGMA index_list(items)").fetchall()
                }
                self.assertIn("idx_items_active_id", indexes)
                self.assertIn("idx_items_favorite_id", indexes)

                library.move_to_recycle_bin(item_id)
                self.assertEqual(library.all_categories(), [])
                self.assertEqual(library.all_source_folders(), [])
                library.restore_from_recycle_bin(item_id)
                self.assertIn("קינוחים/עוגות", library.all_categories())
                library.db.close()
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous

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

    def test_הכנת_ייבוא_ברקע_ושמירה_במסד_בשלב_נפרד(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = str(Path(temp) / "profile")
                source = Path(temp) / "מתכון.txt"
                source.write_text(
                    "עוגת תפוחים\nמצרכים:\n2 כוסות קמח\nאופן ההכנה:\nמערבבים ואופים",
                    encoding="utf-8"
                )
                library = app.RecipeLibrary()
                known = set()
                prepared = library.prepare_import(source, "קינוחים/עוגות", known_digests=known)
                self.assertIn(prepared["digest"], known)
                self.assertEqual(prepared["safe_parts"], ["קינוחים", "עוגות"])
                self.assertIn("2 כוסות קמח", prepared["recipe_data"]["ingredients"])
                added, title = library.import_prepared(prepared)
                self.assertTrue(added)
                self.assertEqual(title, "עוגת תפוחים")
                duplicate = library.prepare_import(source, known_digests=known)
                self.assertTrue(duplicate["duplicate"])
                self.assertFalse(library.import_prepared(duplicate)[0])
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

    def test_גודל_מטמון_הגלריה_מותאם_לספרייה_ומוגבל_בזיכרון(self):
        instance = app.App.__new__(app.App)
        instance._gallery_image_cache = app.OrderedDict(
            (index, object()) for index in range(10)
        )

        self.assertEqual(instance._set_gallery_image_cache_limit(7), 7)
        self.assertEqual(len(instance._gallery_image_cache), 7)

        large_library_limit = instance._set_gallery_image_cache_limit(5000)
        self.assertEqual(large_library_limit, app.GALLERY_CACHE_MAX_IMAGES)
        self.assertEqual(large_library_limit, 542)  # 96 MiB / (320 × 145 × 4 bytes)
        self.assertGreater(large_library_limit, 64)
        self.assertLessEqual(
            large_library_limit * app.GALLERY_THUMBNAIL_ESTIMATED_BYTES,
            app.GALLERY_CACHE_MEMORY_BUDGET_BYTES,
        )

        self.assertEqual(instance._set_gallery_image_cache_limit(2), 2)
        self.assertEqual(len(instance._gallery_image_cache), 2)


    def test_ספרייה_גדולה_עם_שלושת_אלפים_פריטים(self):
        previous = os.environ.get("LOCALAPPDATA")
        library = None
        started = time.perf_counter()
        try:
            with tempfile.TemporaryDirectory() as temp:
                try:
                    os.environ["LOCALAPPDATA"] = str(Path(temp) / "profile")
                    library = app.RecipeLibrary()
                    records = []
                    for index in range(3000):
                        recipe_data = app.json.dumps(
                            {
                                "category": "בדיקת עומס",
                                "image_path": f"images/recipe-{index}.jpg",
                            },
                            ensure_ascii=False,
                        )
                        records.append(
                            (
                                f"מתכון בדיקה {index}",
                                f"recipe-{index}.txt",
                                str(Path(temp) / f"recipe-{index}.txt"),
                                ".txt",
                                "מצרכים: קמח\\nאופן ההכנה: מערבבים",
                                f"scale-test-digest-{index}",
                                "11/11/2026 00:00",
                                recipe_data,
                            )
                        )

                    library.db.executemany(
                        """
                        INSERT INTO items(
                            title, original_name, stored_path, extension,
                            content, digest, imported_at, recipe_data
                        ) VALUES(?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        records,
                    )
                    tag_names = [f"תגית {index}" for index in range(5)]
                    library.db.executemany(
                        "INSERT OR IGNORE INTO tags(name) VALUES(?)",
                        [(name,) for name in tag_names],
                    )
                    tag_ids = {
                        row["name"]: int(row["id"])
                        for row in library.db.execute("SELECT id, name FROM tags")
                    }
                    item_ids = [
                        int(row["id"])
                        for row in library.db.execute("SELECT id FROM items ORDER BY id")
                    ]
                    library.db.executemany(
                        "INSERT INTO item_tags(item_id, tag_id) VALUES(?, ?)",
                        [
                            (item_id, tag_ids[tag_names[index % len(tag_names)]])
                            for index, item_id in enumerate(item_ids)
                        ],
                    )
                    library.db.commit()

                    rows = library.rows()
                    self.assertEqual(len(rows), 3000)
                    tags_by_item = library.tags_by_item(row["id"] for row in rows)
                    self.assertEqual(len(tags_by_item), 3000)
                    self.assertTrue(all(len(names) == 1 for names in tags_by_item.values()))

                    image_count = sum(
                        1 for row in rows
                        if app._recipe_data_from_row(row).get("image_path")
                    )
                    self.assertEqual(image_count, 3000)

                    instance = app.App.__new__(app.App)
                    instance._gallery_image_cache = app.OrderedDict()
                    self.assertEqual(instance._set_gallery_image_cache_limit(image_count), 542)

                    elapsed = time.perf_counter() - started
                    print(
                        f"3,000-item database smoke test passed "
                        f"(batched tag lookup and image count: {elapsed:.3f}s)."
                    )
                finally:
                    if library is not None:
                        library.db.close()
                        library = None
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous


    def test_מטמון_תמונות_הגלריה_מתעדכן_ונשאר_מוגבל(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            source = folder / "תמונה.png"
            Image.new("RGB", (40, 30), "red").save(source)

            instance = app.App.__new__(app.App)
            instance._gallery_image_cache = app.OrderedDict()
            instance._gallery_image_cache_limit = 2

            with patch.object(app.ImageTk, "PhotoImage", side_effect=lambda *_args, **_kwargs: object()) as factory:
                first = instance._gallery_photo(str(source))
                reused = instance._gallery_photo(str(source))
                self.assertIs(first, reused)
                self.assertEqual(factory.call_count, 1)

                old_stat = source.stat()
                Image.new("RGB", (40, 30), "blue").save(source)
                current_stat = source.stat()
                os.utime(
                    source,
                    ns=(current_stat.st_atime_ns, max(current_stat.st_mtime_ns, old_stat.st_mtime_ns) + 2_000_000_000),
                )
                refreshed = instance._gallery_photo(str(source))
                self.assertIsNot(first, refreshed)
                self.assertEqual(len(instance._gallery_image_cache), 1)

                second_path = folder / "תמונה-2.png"
                third_path = folder / "תמונה-3.png"
                Image.new("RGB", (40, 30), "green").save(second_path)
                Image.new("RGB", (40, 30), "yellow").save(third_path)
                second = instance._gallery_photo(str(second_path))
                self.assertIs(second, instance._gallery_photo(str(second_path)))
                self.assertIs(refreshed, instance._gallery_photo(str(source)))
                instance._gallery_photo(str(third_path))

                self.assertLessEqual(len(instance._gallery_image_cache), 2)
                second_normalized = os.path.normcase(str(second_path.resolve()))
                self.assertNotIn(
                    second_normalized,
                    {key[0] for key in instance._gallery_image_cache},
                    "Least-recently-used thumbnails should be evicted first",
                )
                self.assertEqual(factory.call_count, 4)


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
                    self.assertFalse(error, error)
                    self.assertIsNotNone(pdf_path)
                    self.assertTrue(pdf_path.is_file())
                    with app.fitz.open(pdf_path) as pdf:
                        self.assertGreater(pdf.page_count, 0)
                    shutil.rmtree(temp_dir, ignore_errors=True)

    def test_קובץ_office_פגום_לא_מפיל_את_ההמרה(self):
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / "קובץ.docx"
            source.write_bytes(b"this is not a valid Office document")
            pdf_path, temp_dir, error = app.convert_office_to_pdf(source)
            if pdf_path is not None:
                self.assertTrue(pdf_path.is_file())
                with app.fitz.open(pdf_path) as converted:
                    self.assertGreater(converted.page_count, 0)
            else:
                self.assertTrue(error)
            if temp_dir is not None:
                shutil.rmtree(temp_dir, ignore_errors=True)

    def test_שם_ללא_תוכן_נשאר_קריא(self):
        self.assertGreater(app.relevance("קובץ תמונה", "", "תמונה"), 0)


    def test_מועדפים_ואוספים_נשמרים_בין_הפעלות(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = str(Path(temp) / "profile")
                source = Path(temp) / "עוגה.txt"
                source.write_text("עוגת תפוחים\nתפוחים, קמח וסוכר", encoding="utf-8")

                library = app.RecipeLibrary()
                added, title = library.import_one(source)
                self.assertTrue(added)
                item_id = int(library.rows()[0]["id"])

                desserts_id = library.create_collection("  קינוחים  ")
                shabbat_id = library.create_collection("שבת")
                library.set_favorite(item_id, True)
                library.set_item_collection(item_id, desserts_id, True)
                library.set_item_collection(item_id, shabbat_id, True)

                self.assertEqual(library.item_collection_ids(item_id), {desserts_id, shabbat_id})
                self.assertEqual(library.item_collection_names(item_id), ["קינוחים", "שבת"])
                self.assertEqual(len(library.rows(favorite_only=True)), 1)
                self.assertEqual(len(library.rows(collection_id=desserts_id)), 1)
                with self.assertRaises(ValueError):
                    library.create_collection("קינוחים")

                library.db.close()
                reopened = app.RecipeLibrary()
                row = reopened.rows()[0]
                self.assertEqual(row["title"], title)
                self.assertEqual(row["is_favorite"], 1)
                self.assertEqual(reopened.item_collection_ids(item_id), {desserts_id, shabbat_id})

                reopened.set_item_collection(item_id, shabbat_id, False)
                self.assertEqual(reopened.item_collection_names(item_id), ["קינוחים"])
                reopened.delete_collection(desserts_id)
                self.assertEqual(reopened.item_collection_ids(item_id), set())
                self.assertEqual(len(reopened.rows()), 1)
                reopened.db.close()
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous

    def test_שדרוג_מסד_נתונים_ישן_שומר_את_הקבצים(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = temp
                home = Path(temp) / app.APP_NAME
                home.mkdir(parents=True)
                (home / "הקבצים שלי").mkdir()
                db_path = home / "הספרייה.sqlite3"
                old_db = app.sqlite3.connect(db_path)
                old_db.execute("""
                    CREATE TABLE items(
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        title TEXT NOT NULL,
                        original_name TEXT NOT NULL,
                        stored_path TEXT NOT NULL,
                        extension TEXT NOT NULL,
                        content TEXT NOT NULL DEFAULT '',
                        digest TEXT NOT NULL UNIQUE,
                        imported_at TEXT NOT NULL
                    )
                """)
                old_db.execute(
                    "INSERT INTO items(title,original_name,stored_path,extension,content,digest,imported_at) VALUES(?,?,?,?,?,?,?)",
                    ("מתכון ישן", "old.txt", "old.txt", ".txt", "תוכן", "legacy-digest", "01/01/2025 10:00")
                )
                old_db.commit()
                old_db.close()

                library = app.RecipeLibrary()
                self.assertEqual(len(library.rows()), 1)
                self.assertEqual(library.rows()[0]["title"], "מתכון ישן")
                self.assertEqual(library.rows()[0]["is_favorite"], 0)
                self.assertEqual(library.collections(), [])
                library.set_favorite(1, True)
                self.assertEqual(library.rows(favorite_only=True)[0]["id"], 1)
                library.db.close()
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous


    def test_זיהוי_מצרכים_וחישוב_כמויות(self):
        source = "מצרכים:\n2 כוסות קמח\n1/2 כוס סוכר\n3 ביצים\nאופן ההכנה:\nמערבבים היטב"
        extracted = app.extract_ingredients(source)
        self.assertIn("2 כוסות קמח", extracted)
        self.assertIn("1/2 כוס סוכר", extracted)
        self.assertNotIn("מערבבים היטב", extracted)
        scaled = app.scale_ingredient_amounts(extracted, 4, 8)
        self.assertIn("4 כוסות קמח", scaled)
        self.assertIn("1 כוס סוכר", scaled)
        self.assertIn("6 ביצים", scaled)

    def test_נתוני_מתכון_תגיות_וחיפוש_קטגוריות(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = temp
                source = Path(temp) / "עוגה.txt"
                source.write_text("עוגה מיוחדת\n2 כוסות קמח", encoding="utf-8")
                library = app.RecipeLibrary()
                added, _ = library.import_one(source)
                self.assertTrue(added)
                item_id = library.rows()[0]["id"]
                library.set_recipe_data(item_id, {
                    "category": "קינוחים",
                    "ingredients": "2 כוסות קמח",
                    "instructions": "מערבבים ואופים",
                    "prep_minutes": "15",
                    "cook_minutes": "30",
                    "servings": "4",
                    "notes": "לשמור בקירור"
                })
                library.set_item_tags(item_id, ["קל", "מהיר", "קל"])
                self.assertEqual(library.get_recipe_data(item_id)["servings"], "4")
                self.assertEqual(library.tag_names(item_id), ["מהיר", "קל"])
                self.assertIn("קינוחים", library.all_categories())
                self.assertIn("מהיר", library.all_tags())
                library.db.close()
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous

    def test_סל_מחזור_משחזר_קובץ_בלי_לאבד_מטאדאטה(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = temp
                source = Path(temp) / "מסמך.txt"
                source.write_text("מסמך בדיקה", encoding="utf-8")
                library = app.RecipeLibrary()
                added, _ = library.import_one(source)
                self.assertTrue(added)
                row = library.rows()[0]
                item_id = int(row["id"])
                stored = Path(row["stored_path"])
                library.set_favorite(item_id, True)
                library.move_to_recycle_bin(item_id)
                self.assertEqual(library.rows(), [])
                self.assertEqual(len(library.recycled_rows()), 1)
                self.assertTrue(stored.is_file())
                library.restore_from_recycle_bin(item_id)
                restored = library.rows()[0]
                self.assertEqual(restored["is_favorite"], 1)
                self.assertTrue(Path(restored["stored_path"]).is_file())
                library.db.close()
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous

    def test_גיבוי_ושחזור_מחזירים_את_הספרייה_כולה(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = str(Path(temp) / "profile")
                first = Path(temp) / "ראשון.txt"
                first.write_text("מתכון ראשון", encoding="utf-8")
                second = Path(temp) / "שני.txt"
                second.write_text("מתכון שני", encoding="utf-8")

                library = app.RecipeLibrary()
                self.assertTrue(library.import_one(first)[0])
                item_id = int(library.rows()[0]["id"])
                collection_id = library.create_collection("שמורים")
                library.set_favorite(item_id, True)
                library.set_item_collection(item_id, collection_id, True)
                backup = Path(temp) / "גיבוי.zip"
                library.create_backup(backup)

                self.assertTrue(library.import_one(second)[0])
                self.assertEqual(len(library.rows()), 2)
                restored_count = library.restore_backup(backup)
                self.assertEqual(restored_count, 1)
                restored = library.rows()
                self.assertEqual(len(restored), 1)
                self.assertEqual(restored[0]["is_favorite"], 1)
                self.assertTrue(Path(restored[0]["stored_path"]).is_file())
                self.assertEqual(library.item_collection_names(item_id), ["שמורים"])
                library.db.close()
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous

    def test_בדיקת_דמיון_תוכן_לזיהוי_כפילויות(self):
        first = "עוגת תפוחים עם קינמון: קמח, ביצים, תפוחים וסוכר. אופים בתנור."
        second = "עוגת תפוחים קינמון: מערבבים קמח, ביצים, תפוחים וסוכר ואופים בתנור."
        self.assertGreater(app.text_similarity(first, second), 0.55)


    def test_חיפוש_כולל_מצרכים_והוראות_הכנה(self):
        blob = app.recipe_search_text(
            "עוגה פשוטה", "",
            {"ingredients": "2 כוסות קמח", "instructions": "מערבבים ואופים"},
            ["חלבי"]
        )
        self.assertGreater(app.relevance("עוגה פשוטה", blob, "כוסות קמח"), 0)
        self.assertGreater(app.relevance("עוגה פשוטה", blob, "מערבבים"), 0)

    def test_רשימת_קניות_מאחדת_רק_שורות_זהות(self):
        result = app.compile_shopping_list([
            {"ingredients": "2 כוסות קמח\n1 כף מלח"},
            {"ingredients": "2 כוסות קמח\n2 כפות מלח"},
        ])
        self.assertEqual(result.count("2 כוסות קמח"), 1)
        self.assertIn("1 כף מלח", result)
        self.assertIn("2 כפות מלח", result)

    def test_יצירת_חוברת_word_בעברית_ובכיווניות_rtl(self):
        if app.Document is None:
            self.skipTest("נדרשת ספריית python-docx")
        with tempfile.TemporaryDirectory() as temp:
            target = Path(temp) / "חוברת.docx"
            app.build_recipe_book_docx([
                {
                    "title": "עוגת תפוחים",
                    "ingredients": "2 כוסות קמח\n3 תפוחים",
                    "instructions": "",
                    "content": "עוגת תפוחים\n2 כוסות קמח\nאופן ההכנה:\nמערבבים את החומרים\nאופים בתנור",
                    "notes": "להגיש קר",
                    "servings": "8",
                }
            ], "המתכונים שלי", target)
            self.assertTrue(target.is_file())
            document = app.Document(str(target))
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)
            self.assertIn("עוגת תפוחים", text)
            self.assertIn("2 כוסות קמח", text)
            self.assertIn("טקסט נוסף מהקובץ המקורי", text)
            self.assertIn("מערבבים את החומרים", text)
            self.assertIn("הערות אישיות", text)
            self.assertIn("w:bidi", document.paragraphs[0]._p.xml)

    @unittest.skipUnless(
        app.find_libreoffice() is not None and app.fitz is not None,
        "נדרשים LibreOffice ו-PyMuPDF (כלולים בחבילת Windows המלאה)"
    )
    def test_יצירת_חוברת_pdf_עם_כמה_מתכונים(self):
        with tempfile.TemporaryDirectory() as temp:
            target = Path(temp) / "חוברת.pdf"
            recipes = [
                {
                    "title": "עוגת תפוחים",
                    "ingredients": "2 כוסות קמח\n3 תפוחים",
                    "instructions": "מערבבים\nאופים",
                },
                {
                    "title": "מרק ירקות",
                    "ingredients": "2 גזרים\n1 תפוח אדמה",
                    "instructions": "חותכים\nמבשלים",
                },
            ]
            result = app.export_recipe_book_pdf(
                recipes, "החוברת המשפחתית", target, app.convert_office_to_pdf
            )
            self.assertTrue(result.is_file())
            self.assertGreater(result.stat().st_size, 100)
            with app.fitz.open(result) as pdf:
                self.assertGreaterEqual(pdf.page_count, 2)
                extracted = "\n".join(page.get_text() for page in pdf)
                self.assertIn("עוגת תפוחים", extracted)
                self.assertIn("מרק ירקות", extracted)


    def test_ייבוא_תיקיות_שומר_היררכיה_ונתיב_מקורי(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = temp
                source = Path(temp) / "מקור" / "קינוחים" / "עוגות" / "עוגת תפוחים.txt"
                source.parent.mkdir(parents=True)
                source.write_text(
                    "עוגת תפוחים\nמצרכים:\n2 כוסות קמח\n3 תפוחים\nאופן ההכנה:\nמערבבים\nאופים",
                    encoding="utf-8"
                )
                library = app.RecipeLibrary()
                added, title = library.import_one(source, relative_folder="קינוחים/עוגות")
                self.assertTrue(added)
                row = library.rows()[0]
                stored = Path(row["stored_path"])
                self.assertEqual(stored.parent, library.files_dir / "קינוחים" / "עוגות")
                self.assertTrue(stored.is_file())
                data = library.get_recipe_data(row["id"])
                self.assertEqual(data["source_folder"], "קינוחים/עוגות")
                self.assertEqual(data["category"], "קינוחים/עוגות")
                self.assertIn("קינוחים", library.all_source_folders())
                self.assertIn("קינוחים/עוגות", library.all_source_folders())
                self.assertTrue(data["is_recipe"])
                library.db.close()
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous

    def test_ייבוא_חכם_מזהה_מצרכים_והוראות(self):
        parsed = app.parse_recipe_text(
            "עוגת שוקולד\nמצרכים:\n2 כוסות קמח\n1 כוס סוכר\nאופן ההכנה:\nמערבבים היטב\nאופים 30 דקות"
        )
        self.assertEqual(parsed["title"], "עוגת שוקולד")
        self.assertIn("2 כוסות קמח", parsed["ingredients"])
        self.assertIn("מערבבים היטב", parsed["instructions"])
        self.assertTrue(parsed["is_recipe"])
        self.assertEqual(app.recipe_missing_parts(parsed), [])

    def test_זיהוי_מתכונים_חסרים_ומסמכים_רגילים(self):
        self.assertTrue(app.is_recipe_record({"is_recipe": True}, ""))
        self.assertFalse(app.is_recipe_record({"is_recipe": False, "ingredients": "קמח"}, ""))
        self.assertEqual(app.recipe_missing_parts({"ingredients": "קמח"}), ["הוראות הכנה"])
        self.assertEqual(app.recipe_missing_parts({"instructions": "מערבבים"}), ["מצרכים"])

    def test_שמירת_דירוג_היסטוריה_וגריסאות_לא_נמחקת_בעדכון_רגיל(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = temp
                source = Path(temp) / "עוגה.txt"
                source.write_text("עוגה", encoding="utf-8")
                library = app.RecipeLibrary()
                self.assertTrue(library.import_one(source)[0])
                item_id = int(library.rows()[0]["id"])
                library.set_recipe_data(item_id, {
                    "rating": "5",
                    "is_recipe": True,
                    "history": [{"date": "01/01/2026", "loved": True, "changes": "פחות סוכר"}],
                    "versions": [{"label": "כמות כפולה", "ingredients": "4 כוסות קמח"}],
                })
                library.set_recipe_data(item_id, {"category": "קינוחים"})
                data = library.get_recipe_data(item_id)
                self.assertEqual(data["category"], "קינוחים")
                self.assertEqual(data["rating"], "5")
                self.assertEqual(data["history"][0]["changes"], "פחות סוכר")
                self.assertEqual(data["versions"][0]["label"], "כמות כפולה")
                library.db.close()
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous

    def test_תוכן_עניינים_ומספור_עמודים_בחוברת(self):
        if app.Document is None:
            self.skipTest("נדרשת ספריית python-docx")
        with tempfile.TemporaryDirectory() as temp:
            target = Path(temp) / "עיצוב.docx"
            app.build_recipe_book_docx([
                {"title": "מרק ירקות", "ingredients": "2 גזרים", "instructions": "מבשלים"},
                {"title": "עוגת גבינה", "ingredients": "גבינה", "instructions": "אופים"},
            ], "ספר המתכונים", target, template="אלגנטי", layout="כרטיסים קטנים (2 בעמוד)")
            document = app.Document(str(target))
            text = "\n".join(p.text for p in document.paragraphs)
            self.assertIn("תוכן החוברת", text)
            self.assertIn("מרק ירקות", text)
            self.assertIn("עוגת גבינה", text)
            self.assertIn("PAGE", document.sections[0].footer._element.xml)


    def test_שינוי_שם_שומר_את_תיקיית_המקור(self):
        previous = os.environ.get("LOCALAPPDATA")
        try:
            with tempfile.TemporaryDirectory() as temp:
                os.environ["LOCALAPPDATA"] = temp
                source = Path(temp) / "מקור" / "קינוחים" / "עוגות" / "עוגה.txt"
                source.parent.mkdir(parents=True)
                source.write_text("עוגה\nמצרכים:\nקמח\nאופן ההכנה:\nמערבבים", encoding="utf-8")
                library = app.RecipeLibrary()
                self.assertTrue(library.import_one(source, relative_folder="קינוחים/עוגות")[0])
                row = library.rows()[0]
                library.rename(int(row["id"]), "עוגה חדשה")
                renamed = library.rows()[0]
                self.assertEqual(Path(renamed["stored_path"]).parent, library.files_dir / "קינוחים" / "עוגות")
                self.assertEqual(Path(renamed["stored_path"]).name, "עוגה חדשה.txt")
                self.assertTrue(Path(renamed["stored_path"]).is_file())
                self.assertEqual(library.get_recipe_data(row["id"])["source_folder"], "קינוחים/עוגות")
                library.db.close()
        finally:
            if previous is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = previous

if __name__ == "__main__":
    unittest.main()
