# -*- coding: utf-8 -*-
"""
ספריית המתכונים שלי — יישום מקומי בעברית לניהול מתכונים, מסמכים ותמונות.
כל הנתונים נשמרים במחשב בלבד. אין חיבור לרשת ואין שליחת מידע.
"""
from __future__ import annotations

import os
import re
import sys
import csv
import json
import hashlib
import shutil
import sqlite3
import tempfile
import zipfile
import subprocess
import webbrowser
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from html.parser import HTMLParser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from recipe_center import RecipeCenter, recipe_search_text, compile_shopping_list, build_recipe_book_docx, export_recipe_book_pdf, parse_recipe_text, is_recipe_record, recipe_missing_parts

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    ROOT_CLASS = TkinterDnD.Tk
except Exception:
    ROOT_CLASS = tk.Tk
    DND_FILES = None

try:
    import fitz
except Exception:
    fitz = None
try:
    from docx import Document
except Exception:
    Document = None
try:
    from striprtf.striprtf import rtf_to_text
except Exception:
    rtf_to_text = None
try:
    from PIL import Image, ImageTk, ExifTags
except Exception:
    Image = ImageTk = ExifTags = None
try:
    import pytesseract
except Exception:
    pytesseract = None

APP_NAME = "ספריית המתכונים שלי"
APP_VERSION = "1.6.0"
IMAGE_TYPES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"}
TEXT_TYPES = {".txt", ".md", ".csv", ".tsv", ".log"}
WORD_TYPES = {".doc", ".docx", ".docm", ".odt", ".rtf"}
EXCEL_TYPES = {".xls", ".xlsx", ".xlsm", ".ods"}
PRESENTATION_TYPES = {".ppt", ".pptx", ".pptm", ".odp"}
OFFICE_TYPES = WORD_TYPES | EXCEL_TYPES | PRESENTATION_TYPES
EBOOK_TYPES = {".epub", ".mobi", ".fb2"}
DOCUMENT_TYPES = {".pdf", ".html", ".htm"} | OFFICE_TYPES | EBOOK_TYPES
MAX_FILE_SIZE = 250 * 1024 * 1024
MAX_TEXT_LENGTH = 400_000

def find_libreoffice() -> Path | None:
    """Find the bundled or system LibreOffice command-line executable."""
    names = ("soffice.com", "soffice.exe")
    roots = []

    if getattr(sys, "frozen", False):
        roots.append(Path(sys.executable).resolve().parent)
    module_dir = Path(__file__).resolve().parent
    roots.extend((module_dir, module_dir.parent))

    program_files = os.environ.get("PROGRAMFILES")
    if program_files:
        roots.append(Path(program_files))
    program_files_x86 = os.environ.get("PROGRAMFILES(X86)")
    if program_files_x86:
        roots.append(Path(program_files_x86))

    for root in roots:
        for name in names:
            candidate = root / "LibreOffice" / "program" / name
            if candidate.is_file():
                return candidate
            candidate = root / "program" / name
            if candidate.is_file() and candidate.parent.name.lower() == "program":
                return candidate

    for name in ("soffice.com", "soffice.exe", "soffice"):
        located = shutil.which(name)
        if located:
            return Path(located)
    return None


def convert_office_to_pdf(source: Path, timeout: int = 180) -> tuple[Path | None, Path | None, str]:
    """Convert an Office document to a temporary PDF without showing LibreOffice's UI.

    Returns (pdf_path, temporary_directory, error_message). The caller owns
    the temporary directory and must remove it when the preview window closes.
    """
    source = Path(source)
    suffix = source.suffix.lower()
    if suffix not in OFFICE_TYPES:
        return None, None, "סוג הקובץ אינו נתמך להמרה."
    if not source.is_file():
        return None, None, "הקובץ לא נמצא."
    executable = find_libreoffice()
    if executable is None:
        return None, None, "מנוע התצוגה של Office לא נמצא. יש לוודא שתיקיית LibreOffice נמצאת לצד קובץ התוכנה."

    temporary_directory = Path(tempfile.mkdtemp(prefix="ספריית-מתכונים-office-"))
    output_directory = temporary_directory / "output"
    profile_directory = temporary_directory / "profile"
    output_directory.mkdir()
    profile_directory.mkdir()
    output_pdf = output_directory / (source.stem + ".pdf")
    command = [
        str(executable),
        "--headless",
        "--nologo",
        "--nodefault",
        "--nolockcheck",
        f"-env:UserInstallation={profile_directory.resolve().as_uri()}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_directory),
        str(source.resolve()),
    ]
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=max(10, int(timeout)),
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if not output_pdf.is_file() or output_pdf.stat().st_size < 100:
            detail = (completed.stderr or completed.stdout or "").strip()
            detail = detail[-700:] if detail else "LibreOffice לא הפיק קובץ PDF תקין."
            raise RuntimeError(detail)
        if fitz is not None:
            with fitz.open(output_pdf) as converted:
                if converted.page_count < 1:
                    raise RuntimeError("קובץ ה-PDF שהתקבל אינו מכיל עמודים.")
        return output_pdf, temporary_directory, ""
    except subprocess.TimeoutExpired:
        shutil.rmtree(temporary_directory, ignore_errors=True)
        return None, None, "המרת הקובץ ארכה יותר מדי זמן והופסקה כדי למנוע תקיעה."
    except Exception as exc:
        shutil.rmtree(temporary_directory, ignore_errors=True)
        return None, None, f"לא ניתן להציג את מסמך Office: {exc}"


def configure_ocr():
    """Locate the bundled offline Tesseract engine and Hebrew language data."""
    if pytesseract is None:
        return False
    module_dir = Path(__file__).resolve().parent
    base = Path(getattr(sys, "_MEIPASS", module_dir))
    candidates = [
        base / "tesseract" / "tesseract.exe",
        base / "tesseract.exe",
        module_dir.parent / "tesseract" / "tesseract.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            pytesseract.pytesseract.tesseract_cmd = str(candidate)
            tessdata = candidate.parent / "tessdata"
            if tessdata.exists():
                os.environ["TESSDATA_PREFIX"] = str(tessdata)
            return True
    return False

def ocr_image(image) -> str:
    if pytesseract is None:
        return ""
    try:
        configure_ocr()
        return pytesseract.image_to_string(image, lang="heb+eng", config="--psm 6", timeout=60)
    except Exception:
        try:
            return pytesseract.image_to_string(image, lang="eng", config="--psm 6", timeout=60)
        except Exception:
            return ""

def data_home() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    folder = Path(base) / APP_NAME
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "הקבצים שלי").mkdir(exist_ok=True)
    return folder

def normalize(value: str) -> str:
    value = (value or "").lower()
    value = re.sub(r"[\u0591-\u05C7]", "", value)
    value = re.sub(r"[^0-9a-zא-ת]+", " ", value)
    return " ".join(value.split())

def safe_title(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", value or "")
    value = re.sub(r"\s+", " ", value).strip(" .")
    return (value or "מסמך ללא כותרת")[:120]

def unique_path(folder: Path, title: str, suffix: str) -> Path:
    base = safe_title(title)
    candidate = folder / (base + suffix)
    number = 2
    while candidate.exists():
        candidate = folder / (base + f" ({number})" + suffix)
        number += 1
    return candidate

def decode_text(raw: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-16", "cp1255", "cp1252", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", errors="replace")

class TextCollector(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []
        self.hidden = 0
    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.hidden += 1
    def handle_endtag(self, tag):
        if tag in ("script", "style") and self.hidden:
            self.hidden -= 1
        if tag in ("p", "div", "br", "li", "h1", "h2", "h3", "tr"):
            self.parts.append("\n")
    def handle_data(self, data):
        if not self.hidden and data.strip():
            self.parts.append(data.strip() + " ")
    def text(self):
        return "\n".join(self.parts)

def office_xml_text(path: Path, suffix: str) -> str:
    pieces = []
    with zipfile.ZipFile(path) as archive:
        if suffix in (".docx", ".docm"):
            names = ["word/document.xml"]
        elif suffix == ".odt":
            names = ["content.xml"]
        elif suffix in (".xlsx", ".xlsm"):
            names = [n for n in archive.namelist() if n == "xl/sharedStrings.xml" or re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)]
        elif suffix in (".pptx", ".pptm"):
            names = sorted(n for n in archive.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
        else:
            names = []
        for name in names:
            try:
                root = ET.fromstring(archive.read(name))
                for node in root.iter():
                    if node.tag.rsplit("}", 1)[-1] in ("p", "t", "v"):
                        if node.text and node.text.strip():
                            pieces.append(node.text.strip())
            except Exception:
                continue
    return "\n".join(pieces)

def image_description(path: Path) -> str:
    if Image is None:
        return ""
    try:
        with Image.open(path) as image:
            info = image.info or {}
            for key in ("Description", "Title", "Comment", "caption"):
                value = info.get(key)
                if isinstance(value, str) and value.strip():
                    return value.strip()
            exif = image.getexif()
            if exif:
                for key, value in exif.items():
                    name = ExifTags.TAGS.get(key, "")
                    if name in ("ImageDescription", "DocumentName"):
                        if isinstance(value, bytes):
                            value = value.decode("utf-8", errors="ignore")
                        if str(value).strip():
                            return str(value).strip()
                    if name == "XPTitle" and isinstance(value, bytes):
                        text = value.decode("utf-16-le", errors="ignore").strip("\x00 ")
                        if text:
                            return text
    except Exception:
        pass
    return ""

def extract_text(path: Path) -> str:
    suffix = path.suffix.lower()
    try:
        if suffix in TEXT_TYPES:
            text = decode_text(path.read_bytes())
        elif suffix in (".html", ".htm"):
            parser = TextCollector()
            parser.feed(decode_text(path.read_bytes()))
            text = parser.text()
        elif suffix == ".pdf" and fitz is not None:
            with fitz.open(path) as pdf:
                pages = []
                for page in pdf:
                    page_text = page.get_text("text").strip()
                    if not page_text and Image is not None:
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                        import io
                        page_text = ocr_image(Image.open(io.BytesIO(pix.tobytes("png"))))
                    pages.append(page_text)
                text = "\n".join(pages)
        elif suffix in EBOOK_TYPES and fitz is not None:
            with fitz.open(path) as ebook:
                text = "\n".join(page.get_text("text").strip() for page in ebook)
        elif suffix in (".docx", ".docm") and Document is not None:
            doc = Document(path)
            parts = [p.text for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    parts.append(" | ".join(cell.text for cell in row.cells))
            text = "\n".join(parts)
        elif suffix == ".rtf" and rtf_to_text is not None:
            text = rtf_to_text(decode_text(path.read_bytes()))
        elif suffix in (".odt", ".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm"):
            text = office_xml_text(path, suffix)
        elif suffix in IMAGE_TYPES:
            metadata = image_description(path)
            recognized = ""
            if Image is not None:
                with Image.open(path) as image:
                    recognized = ocr_image(image.convert("RGB"))
            text = "\\n".join(part for part in (metadata, recognized) if part)
        else:
            text = ""
    except Exception:
        text = ""
    return (text or "")[:MAX_TEXT_LENGTH]

def choose_title(path: Path, content: str) -> str:
    lines = [re.sub(r"\s+", " ", line).strip(" \t#*-•") for line in (content or "").splitlines()]
    for line in lines[:35]:
        line = re.sub(r"^(?:כותרת|שם המתכון|מתכון)\s*[:：-]\s*", "", line, flags=re.I)
        if 2 <= len(line) <= 120 and len(line.split()) <= 16:
            if not re.fullmatch(r"[\d\W_]+", line):
                return safe_title(line)
    metadata = image_description(path) if path.suffix.lower() in IMAGE_TYPES else ""
    return safe_title(metadata or path.stem)

def file_kind(suffix: str) -> str:
    suffix = (suffix or "").lower()
    if suffix in IMAGE_TYPES:
        return "תמונה"
    if suffix in EXCEL_TYPES:
        return "גיליון"
    if suffix in PRESENTATION_TYPES:
        return "מצגת"
    if suffix in EBOOK_TYPES:
        return "ספר דיגיטלי"
    if suffix == ".zip":
        return "ארכיון"
    if suffix in TEXT_TYPES or suffix in DOCUMENT_TYPES:
        return "מסמך"
    return "קובץ"


def file_digest(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()

def relevance(title: str, content: str, query: str) -> int:
    q = normalize(query)
    if not q:
        return 1
    name = normalize(title)
    body = normalize(content)
    if q in name:
        return 1200 - max(0, len(name) - len(q))
    words = q.split()
    hits = sum(1 for word in words if word in name)
    if hits == len(words):
        return 950
    if hits:
        return 600 + hits * 20
    if q in body:
        return 250
    body_hits = sum(1 for word in words if word in body)
    if words and body_hits == len(words):
        return 180
    ratio = SequenceMatcher(None, q, name).ratio()
    if len(q) >= 4 and ratio >= 0.78:
        return int(ratio * 100)
    return 0


def extract_ingredients(text: str) -> str:
    """Conservatively collect likely ingredient lines from extracted recipe text."""
    units = ("גרם", "קילו", "ק\"ג", "מ\"ל", "ליטר", "כוס", "כפות", "כף",
             "כפית", "כפיות", "שיניים", "יחידות", "ביצים", "ביצה",
             "חבילה", "שקית", "קורט")
    candidates = []
    for raw in (text or "").splitlines():
        line = re.sub(r"^\s*(?:[-*•▪●]|\d+[.)])\s*", "", raw).strip()
        if not line or len(line) > 180:
            continue
        if any(word in line for word in ("אופן ההכנה", "הוראות הכנה", "שלבי הכנה")):
            continue
        starts_with_quantity = bool(re.match(r"^(?:\d+(?:[.,]\d+)?(?:\s*/\s*\d+)?|[½¼¾⅓⅔⅛⅜⅝⅞])\s*", line))
        has_unit = any(unit in line for unit in units)
        if starts_with_quantity or has_unit:
            candidates.append(line)
        if len(candidates) >= 60:
            break
    return "\n".join(candidates)


def scale_ingredient_amounts(text: str, base_servings: int, target_servings: int) -> str:
    base = max(1, int(base_servings))
    target = max(1, int(target_servings))
    factor = target / base
    fractions = {"½": 0.5, "¼": 0.25, "¾": 0.75, "⅓": 1/3, "⅔": 2/3,
                 "⅛": 0.125, "⅜": 0.375, "⅝": 0.625, "⅞": 0.875}
    pattern = re.compile(r"^(\s*)(\d+\s*/\s*\d+|\d+(?:[.,]\d+)?|[½¼¾⅓⅔⅛⅜⅝⅞])(\s*)(.*)$")
    output = []
    for line in (text or "").splitlines():
        match = pattern.match(line)
        if not match:
            output.append(line)
            continue
        prefix, raw, spacing, rest = match.groups()
        clean = raw.replace(" ", "").replace(",", ".")
        if clean in fractions:
            amount = fractions[clean]
        elif "/" in clean:
            numerator, denominator = clean.split("/", 1)
            try:
                amount = float(numerator) / float(denominator)
            except (ValueError, ZeroDivisionError):
                output.append(line)
                continue
        else:
            try:
                amount = float(clean)
            except ValueError:
                output.append(line)
                continue
        result = round(amount * factor, 2)
        formatted = str(int(result)) if float(result).is_integer() else f"{result:.2f}".rstrip("0").rstrip(".")
        output.append(f"{prefix}{formatted}{spacing}{rest}")
    return "\n".join(output)


def text_similarity(left: str, right: str) -> float:
    # Token-set overlap is intentionally used here instead of a full edit-distance
    # calculation; duplicate scans can compare thousands of pairs.
    a = normalize((left or "")[:3500])
    b = normalize((right or "")[:3500])
    if not a or not b:
        return 0.0
    a_words, b_words = set(a.split()), set(b.split())
    union = a_words | b_words
    word_score = len(a_words & b_words) / len(union) if union else 0.0
    a_pairs = set(zip(a.split(), a.split()[1:]))
    b_pairs = set(zip(b.split(), b.split()[1:]))
    pair_union = a_pairs | b_pairs
    pair_score = len(a_pairs & b_pairs) / len(pair_union) if pair_union else 0.0
    return max(word_score, pair_score)


class RecipeLibrary:
    def __init__(self):
        self.home = data_home()
        self.files_dir = self.home / "הקבצים שלי"
        self.recycle_dir = self.home / "סל המחזור"
        self.backup_dir = self.home / "גיבויים"
        self.files_dir.mkdir(parents=True, exist_ok=True)
        self.recycle_dir.mkdir(parents=True, exist_ok=True)
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.home / "הספרייה.sqlite3"
        self.db = sqlite3.connect(self.db_path)
        self.db.row_factory = sqlite3.Row
        self._ensure_schema()

    def _ensure_schema(self):
        self.db.execute("PRAGMA foreign_keys = ON")
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS items(
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
        columns = {
            row["name"] for row in self.db.execute("PRAGMA table_info(items)").fetchall()
        }
        migrations = {
            "is_favorite": "INTEGER NOT NULL DEFAULT 0",
            "is_deleted": "INTEGER NOT NULL DEFAULT 0",
            "deleted_at": "TEXT NOT NULL DEFAULT ''",
            "last_opened": "TEXT NOT NULL DEFAULT ''",
            "recipe_data": "TEXT NOT NULL DEFAULT '{}'",
        }
        for name, declaration in migrations.items():
            if name not in columns:
                self.db.execute(f"ALTER TABLE items ADD COLUMN {name} {declaration}")
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS collections(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE COLLATE NOCASE,
                created_at TEXT NOT NULL
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS item_collections(
                item_id INTEGER NOT NULL,
                collection_id INTEGER NOT NULL,
                PRIMARY KEY(item_id, collection_id),
                FOREIGN KEY(item_id) REFERENCES items(id) ON DELETE CASCADE,
                FOREIGN KEY(collection_id) REFERENCES collections(id) ON DELETE CASCADE
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS tags(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE COLLATE NOCASE
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS item_tags(
                item_id INTEGER NOT NULL,
                tag_id INTEGER NOT NULL,
                PRIMARY KEY(item_id, tag_id),
                FOREIGN KEY(item_id) REFERENCES items(id) ON DELETE CASCADE,
                FOREIGN KEY(tag_id) REFERENCES tags(id) ON DELETE CASCADE
            )
        """)
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS app_settings(
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)
        self.db.commit()

    def rows(self, favorite_only: bool = False, collection_id: int | None = None, include_deleted: bool = False):
        conditions = []
        params = []
        if not include_deleted:
            conditions.append("items.is_deleted = 0")
        if favorite_only:
            conditions.append("items.is_favorite = 1")
        if collection_id is not None:
            conditions.append("EXISTS (SELECT 1 FROM item_collections ic WHERE ic.item_id=items.id AND ic.collection_id=?)")
            params.append(int(collection_id))
        where = " WHERE " + " AND ".join(conditions) if conditions else ""
        return self.db.execute(
            "SELECT items.* FROM items" + where + " ORDER BY items.id DESC",
            params
        ).fetchall()

    def recycled_rows(self):
        return self.db.execute(
            "SELECT * FROM items WHERE is_deleted = 1 ORDER BY deleted_at DESC, id DESC"
        ).fetchall()

    def move_to_recycle_bin(self, item_id: int):
        self.db.execute(
            "UPDATE items SET is_deleted=1, deleted_at=? WHERE id=? AND is_deleted=0",
            (datetime.now().strftime("%d/%m/%Y %H:%M"), int(item_id))
        )
        self.db.commit()

    def restore_from_recycle_bin(self, item_id: int):
        self.db.execute(
            "UPDATE items SET is_deleted=0, deleted_at='' WHERE id=? AND is_deleted=1",
            (int(item_id),)
        )
        self.db.commit()

    def permanently_delete(self, item_id: int):
        row = self.db.execute(
            "SELECT stored_path FROM items WHERE id=? AND is_deleted=1", (int(item_id),)
        ).fetchone()
        if row:
            Path(row["stored_path"]).unlink(missing_ok=True)
            self.db.execute("DELETE FROM items WHERE id=?", (int(item_id),))
            self.db.commit()

    def set_favorite(self, item_id: int, is_favorite: bool):
        self.db.execute(
            "UPDATE items SET is_favorite = ? WHERE id = ?",
            (1 if is_favorite else 0, int(item_id))
        )
        self.db.commit()

    def record_opened(self, item_id: int):
        self.db.execute(
            "UPDATE items SET last_opened=? WHERE id=?",
            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), int(item_id))
        )
        self.db.commit()

    def get_recipe_data(self, item_id: int) -> dict:
        row = self.db.execute("SELECT recipe_data FROM items WHERE id=?", (int(item_id),)).fetchone()
        if not row:
            return {}
        try:
            value = json.loads(row["recipe_data"] or "{}")
            return value if isinstance(value, dict) else {}
        except (TypeError, ValueError):
            return {}

    def set_recipe_data(self, item_id: int, data: dict):
        allowed = {
            "category", "ingredients", "instructions", "prep_minutes", "cook_minutes",
            "servings", "notes", "source_folder", "rating", "is_recipe",
            "history", "versions", "image_path"
        }
        existing = self.get_recipe_data(item_id)
        clean = {}
        for key in allowed:
            if key in data:
                clean[key] = data[key]
            elif key in existing:
                clean[key] = existing[key]
            else:
                clean[key] = [] if key in ("history", "versions") else (None if key == "is_recipe" else "")
        self.db.execute(
            "UPDATE items SET recipe_data=? WHERE id=?",
            (json.dumps(clean, ensure_ascii=False), int(item_id))
        )
        self.db.commit()

    def all_source_folders(self) -> list[str]:
        folders = set()
        for row in self.db.execute("SELECT recipe_data FROM items WHERE is_deleted=0"):
            try:
                folder = str(json.loads(row["recipe_data"] or "{}").get("source_folder") or "").strip("/")
            except (ValueError, AttributeError, TypeError):
                folder = ""
            if folder:
                parts = folder.split("/")
                for index in range(1, len(parts) + 1):
                    folders.add("/".join(parts[:index]))
        return sorted(folders, key=lambda value: (value.count("/"), value.casefold()))

    def tag_names(self, item_id: int) -> list[str]:
        rows = self.db.execute("""
            SELECT tags.name FROM tags
            INNER JOIN item_tags ON item_tags.tag_id=tags.id
            WHERE item_tags.item_id=? ORDER BY tags.name COLLATE NOCASE
        """, (int(item_id),)).fetchall()
        return [row["name"] for row in rows]

    def all_tags(self) -> list[str]:
        return [row["name"] for row in self.db.execute("SELECT name FROM tags ORDER BY name COLLATE NOCASE")]

    def all_categories(self) -> list[str]:
        values = set()
        for row in self.db.execute("SELECT recipe_data FROM items WHERE is_deleted=0"):
            try:
                category = json.loads(row["recipe_data"] or "{}").get("category", "").strip()
            except (ValueError, AttributeError):
                category = ""
            if category:
                values.add(category)
        return sorted(values, key=str.casefold)

    def set_item_tags(self, item_id: int, names):
        cleaned = []
        for name in names:
            value = " ".join(str(name).split()).strip()
            if value and value.casefold() not in {x.casefold() for x in cleaned}:
                cleaned.append(value[:32])
        self.db.execute("DELETE FROM item_tags WHERE item_id=?", (int(item_id),))
        for name in cleaned:
            self.db.execute("INSERT OR IGNORE INTO tags(name) VALUES(?)", (name,))
            tag = self.db.execute("SELECT id FROM tags WHERE name=? COLLATE NOCASE", (name,)).fetchone()
            self.db.execute(
                "INSERT OR IGNORE INTO item_tags(item_id,tag_id) VALUES(?,?)",
                (int(item_id), int(tag["id"]))
            )
        self.db.commit()

    def set_setting(self, key: str, value: str):
        self.db.execute(
            "INSERT INTO app_settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (str(key), str(value))
        )
        self.db.commit()

    def get_setting(self, key: str, default: str = "") -> str:
        row = self.db.execute("SELECT value FROM app_settings WHERE key=?", (str(key),)).fetchone()
        return str(row["value"]) if row else str(default)

    def create_backup(self, target: Path):
        target = Path(target)
        target.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="recipe-backup-") as temp:
            db_copy = Path(temp) / "library.sqlite3"
            backup_db = sqlite3.connect(db_copy)
            try:
                self.db.commit()
                self.db.backup(backup_db)
            finally:
                backup_db.close()
            with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.write(db_copy, "הספרייה.sqlite3")
                for path in self.files_dir.rglob("*"):
                    if path.is_file():
                        archive.write(path, "הקבצים שלי/" + path.relative_to(self.files_dir).as_posix())

    def restore_backup(self, source: Path) -> int:
        source = Path(source)
        if not source.is_file() or not zipfile.is_zipfile(source):
            raise ValueError("קובץ הגיבוי אינו ZIP תקין.")
        temp_root = Path(tempfile.mkdtemp(prefix="recipe-restore-", dir=self.home))
        temp_db = temp_root / "הספרייה.sqlite3"
        extracted_files = temp_root / "הקבצים שלי"
        extracted_files.mkdir()
        total_size = 0
        try:
            with zipfile.ZipFile(source, "r") as archive:
                names = archive.namelist()
                if "הספרייה.sqlite3" not in names:
                    raise ValueError("הגיבוי לא מכיל את מסד הנתונים של הספרייה.")
                for info in archive.infolist():
                    name = info.filename.replace("\\", "/")
                    parts = Path(name).parts
                    if name.startswith("/") or ".." in parts or (len(parts) and ":" in parts[0]):
                        raise ValueError("הגיבוי מכיל נתיב קובץ שאינו בטוח.")
                    if name == "הספרייה.sqlite3":
                        destination = temp_db
                    elif name.startswith("הקבצים שלי/"):
                        relative = name[len("הקבצים שלי/"):]
                        parts = [part for part in relative.split("/") if part not in ("", ".")]
                        if not parts:
                            continue
                        if any(part == ".." or ":" in part for part in parts):
                            raise ValueError("הגיבוי מכיל נתיב קובץ שאינו בטוח.")
                        destination = extracted_files.joinpath(*parts)
                        if not destination.resolve().is_relative_to(extracted_files.resolve()):
                            raise ValueError("הגיבוי מכיל נתיב קובץ מחוץ לתיקיית הספרייה.")
                    else:
                        continue
                    if info.is_dir():
                        destination.mkdir(parents=True, exist_ok=True)
                        continue
                    total_size += info.file_size
                    if total_size > 2 * 1024 * 1024 * 1024:
                        raise ValueError("הגיבוי גדול מדי לשחזור.")
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    with archive.open(info) as src, destination.open("wb") as dst:
                        shutil.copyfileobj(src, dst)
            if not temp_db.is_file():
                raise ValueError("מסד הנתונים לא חולץ.")
            check_db = sqlite3.connect(temp_db)
            check_db.row_factory = sqlite3.Row
            integrity = check_db.execute("PRAGMA quick_check").fetchone()[0]
            if integrity != "ok":
                check_db.close()
                raise ValueError("מסד הנתונים שבגיבוי אינו תקין.")
            if not check_db.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='items'").fetchone():
                check_db.close()
                raise ValueError("הקובץ אינו גיבוי של ספריית המתכונים.")
            rows = check_db.execute("SELECT id,stored_path FROM items").fetchall()
            for row in rows:
                filename = Path(str(row["stored_path"])).name
                if filename and (extracted_files / filename).is_file():
                    check_db.execute(
                        "UPDATE items SET stored_path=? WHERE id=?",
                        (str(self.files_dir / filename), int(row["id"]))
                    )
            check_db.commit()
            check_db.close()

            old_files = temp_root / "old_files"
            old_db = temp_root / "old_library.sqlite3"
            self.db.commit()
            self.db.close()
            moved_old_files = False
            moved_old_db = False
            try:
                if self.files_dir.exists():
                    os.replace(self.files_dir, old_files)
                    moved_old_files = True
                if self.db_path.exists():
                    os.replace(self.db_path, old_db)
                    moved_old_db = True
                os.replace(extracted_files, self.files_dir)
                shutil.copy2(temp_db, self.db_path)
                self.db = sqlite3.connect(self.db_path)
                self.db.row_factory = sqlite3.Row
                self._ensure_schema()
            except Exception:
                try:
                    if hasattr(self, "db"):
                        self.db.close()
                except Exception:
                    pass
                if self.files_dir.exists():
                    shutil.rmtree(self.files_dir, ignore_errors=True)
                if self.db_path.exists():
                    self.db_path.unlink(missing_ok=True)
                if moved_old_files:
                    os.replace(old_files, self.files_dir)
                if moved_old_db:
                    os.replace(old_db, self.db_path)
                self.db = sqlite3.connect(self.db_path)
                self.db.row_factory = sqlite3.Row
                self._ensure_schema()
                raise
            return len(rows)
        finally:
            shutil.rmtree(temp_root, ignore_errors=True)

    def set_favorite(self, item_id: int, is_favorite: bool):
        self.db.execute(
            "UPDATE items SET is_favorite = ? WHERE id = ?",
            (1 if is_favorite else 0, int(item_id))
        )
        self.db.commit()

    def create_collection(self, name: str) -> int:
        clean_name = " ".join((name or "").split())
        if not clean_name:
            raise ValueError("צריך להזין שם לאוסף.")
        if len(clean_name) > 48:
            raise ValueError("שם האוסף יכול להכיל עד 48 תווים.")
        try:
            cursor = self.db.execute(
                "INSERT INTO collections(name, created_at) VALUES(?, ?)",
                (clean_name, datetime.now().strftime("%d/%m/%Y %H:%M"))
            )
            self.db.commit()
            return int(cursor.lastrowid)
        except sqlite3.IntegrityError as exc:
            raise ValueError("כבר קיים אוסף בשם הזה.") from exc

    def collections(self):
        return self.db.execute(
            "SELECT * FROM collections ORDER BY name COLLATE NOCASE"
        ).fetchall()

    def delete_collection(self, collection_id: int):
        self.db.execute("DELETE FROM collections WHERE id = ?", (int(collection_id),))
        self.db.commit()

    def item_collection_ids(self, item_id: int) -> set[int]:
        rows = self.db.execute(
            "SELECT collection_id FROM item_collections WHERE item_id = ?",
            (int(item_id),)
        ).fetchall()
        return {int(row["collection_id"]) for row in rows}

    def set_item_collection(self, item_id: int, collection_id: int, assigned: bool):
        if assigned:
            self.db.execute(
                "INSERT OR IGNORE INTO item_collections(item_id, collection_id) VALUES(?, ?)",
                (int(item_id), int(collection_id))
            )
        else:
            self.db.execute(
                "DELETE FROM item_collections WHERE item_id = ? AND collection_id = ?",
                (int(item_id), int(collection_id))
            )
        self.db.commit()

    def item_collection_names(self, item_id: int) -> list[str]:
        rows = self.db.execute("""
            SELECT collections.name
            FROM collections
            INNER JOIN item_collections ON item_collections.collection_id = collections.id
            WHERE item_collections.item_id = ?
            ORDER BY collections.name COLLATE NOCASE
        """, (int(item_id),)).fetchall()
        return [row["name"] for row in rows]

    def import_one(self, source: Path, relative_folder: str = "") -> tuple[bool, str]:
        if not source.is_file():
            return False, "זה אינו קובץ."
        if source.stat().st_size > MAX_FILE_SIZE:
            return False, "הקובץ גדול מדי. ניתן לייבא קבצים עד 250 מגה־בייט."
        digest = file_digest(source)
        if self.db.execute("SELECT id FROM items WHERE digest=?", (digest,)).fetchone():
            return False, "הקובץ כבר נמצא בספרייה."
        content = extract_text(source)
        title = choose_title(source, content)
        safe_parts = []
        for part in Path(str(relative_folder or "")).parts:
            if part in ("", ".", "..", "/", "\\"):
                continue
            cleaned = safe_title(part)
            if cleaned and cleaned not in (".", ".."):
                safe_parts.append(cleaned)
        target_folder = self.files_dir.joinpath(*safe_parts)
        target_folder.mkdir(parents=True, exist_ok=True)
        target = unique_path(target_folder, title, source.suffix.lower())
        parsed = parse_recipe_text(content)
        recipe_data = {
            "category": "/".join(safe_parts),
            "ingredients": parsed.get("ingredients", ""),
            "instructions": parsed.get("instructions", ""),
            "prep_minutes": "",
            "cook_minutes": "",
            "servings": "",
            "notes": "",
            "source_folder": "/".join(safe_parts),
            "rating": "",
            "is_recipe": bool(parsed.get("is_recipe")),
            "history": [],
            "versions": [],
        }
        shutil.copy2(source, target)
        try:
            self.db.execute(
                "INSERT INTO items(title,original_name,stored_path,extension,content,digest,imported_at,recipe_data) VALUES(?,?,?,?,?,?,?,?)",
                (title, source.name, str(target), source.suffix.lower(), content, digest,
                 datetime.now().strftime("%d/%m/%Y %H:%M"), json.dumps(recipe_data, ensure_ascii=False))
            )
            self.db.commit()
        except Exception:
            target.unlink(missing_ok=True)
            raise
        return True, title

    def remove(self, item_id: int):
        row = self.db.execute("SELECT stored_path FROM items WHERE id=?", (item_id,)).fetchone()
        if row:
            Path(row["stored_path"]).unlink(missing_ok=True)
            self.db.execute("DELETE FROM items WHERE id=?", (item_id,))
            self.db.commit()

    def rename(self, item_id: int, title: str):
        row = self.db.execute("SELECT stored_path FROM items WHERE id=?", (item_id,)).fetchone()
        if not row:
            return
        old = Path(row["stored_path"])
        new = unique_path(old.parent, title, old.suffix)
        old.rename(new)
        self.db.execute("UPDATE items SET title=?,stored_path=? WHERE id=?", (safe_title(title), str(new), item_id))
        self.db.commit()

class App:
    def __init__(self):
        self.library = RecipeLibrary()
        self.root = ROOT_CLASS()
        self.root.title(f"{APP_NAME} – גרסה {APP_VERSION}")
        try:
            resource_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
            icon_path = resource_root / "assets" / "app_icon.ico"
            if icon_path.is_file():
                self.root.iconbitmap(default=str(icon_path))
        except (tk.TclError, OSError):
            pass
        self.root.geometry("1400x880")
        self.root.minsize(1080, 700)
        self.style = ttk.Style(self.root)
        self._apply_styles()
        self.search_value = tk.StringVar()
        self.filter_value = tk.StringVar(value="כל הקבצים")
        self.type_filter_value = tk.StringVar(value="כל הסוגים")
        self.category_filter_value = tk.StringVar(value="כל הקטגוריות")
        self.tag_filter_value = tk.StringVar(value="כל התגיות")
        self.sort_value = tk.StringVar(value="תאריך: חדש לישן")
        self.view_mode = tk.StringVar(value=self.library.get_setting("view_mode", "רשימה"))
        if self.view_mode.get() not in ("רשימה", "גלריה"):
            self.view_mode.set("רשימה")
        self.selected_item_id = None
        self.gallery_photos = []
        self.date_from_value = tk.StringVar()
        self.date_to_value = tk.StringVar()
        self.folder_filter_value = tk.StringVar(value="כל התיקיות")
        self.recipe_filter_value = tk.StringVar(value="כל הפריטים")
        self._filter_options = {}
        self.favorite_button = None
        self.preview_temp_directory = None
        self.preview_render_path = None
        self.search_entry = None
        self._auto_backup_job = None
        self.status_value = tk.StringVar(value="ברוכים הבאים! אפשר להוסיף קבצים או לגרור אותם לחלון.")
        self.preview_photo = None
        self.preview_path = None
        self._preview_resize_job = None
        self.current_rows = {}
        self._build()
        self.refresh()
        if DND_FILES is not None:
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self.on_drop)
            except Exception:
                pass
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self._bind_shortcuts()
        self.root.after(2500, self._maybe_auto_backup)

    def _apply_styles(self):
        """Apply a modern, high-contrast desktop theme."""
        theme = self.library.get_setting("color_theme", "כחול")
        palettes = {
            "כחול": ("#EEF3F9", "#FFFFFF", "#172A45", "#10243D", "#25364D", "#718198", "#DCE5EF", "#087E8B", "#066571", "#F5F8FC", "#DCEFF2", "#B83E4B"),
            "ירוק": ("#EFF6F3", "#FFFFFF", "#173E35", "#12362F", "#25443D", "#72877F", "#D8E7E0", "#16846D", "#116B58", "#F4F9F6", "#DDEFE7", "#B83E4B"),
            "סגול": ("#F3F0F8", "#FFFFFF", "#33274D", "#261C3C", "#3D3450", "#80758F", "#E2DCEC", "#7655A6", "#61438D", "#F8F6FB", "#EAE2F5", "#B83E4B"),
        }
        values = palettes.get(theme, palettes["כחול"])
        names = ("background", "surface", "navy", "hero", "text", "muted", "border", "accent", "accent_hover", "row_alt", "selected", "danger")
        self.colors = dict(zip(names, values))
        try:
            font_size = max(8, min(15, int(self.library.get_setting("font_size", "10"))))
        except ValueError:
            font_size = 10
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        c = self.colors
        self.root.configure(background=c["background"])
        self.style.configure(".", font=("Segoe UI", font_size), foreground=c["text"])
        self.style.configure("App.TFrame", background=c["background"])
        self.style.configure("Card.TFrame", background=c["surface"])
        self.style.configure("Panel.TFrame", background=c["surface"])
        self.style.configure("Panel.TLabel", background=c["surface"], foreground=c["text"])
        self.style.configure("Muted.TLabel", background=c["surface"], foreground=c["muted"], font=("Segoe UI", 9))
        self.style.configure("Section.TLabel", background=c["surface"], foreground=c["navy"], font=("Segoe UI", 14, "bold"))
        self.style.configure("Status.TLabel", background=c["surface"], foreground=c["muted"], padding=(12, 8), font=("Segoe UI", 9))

        self.style.configure(
            "Card.TLabelframe", background=c["surface"],
            bordercolor=c["border"], borderwidth=1, relief="solid", padding=14
        )
        self.style.configure(
            "Card.TLabelframe.Label", background=c["surface"],
            foreground=c["navy"], font=("Segoe UI", 12, "bold")
        )
        self.style.configure(
            "TButton", font=("Segoe UI", max(8, font_size - 1), "bold"), padding=(14, 10),
            background="#E8EEF5", foreground=c["navy"],
            bordercolor=c["border"], focuscolor=c["accent"], relief="flat"
        )
        self.style.map(
            "TButton",
            background=[("disabled", "#EEF2F6"), ("pressed", "#D7E2EC"), ("active", "#DCE7F0")],
            foreground=[("disabled", "#9AA7B5"), ("active", c["navy"])]
        )
        self.style.configure(
            "Primary.TButton", background=c["accent"], foreground="#FFFFFF",
            bordercolor=c["accent"], padding=(17, 11)
        )
        self.style.map(
            "Primary.TButton",
            background=[("disabled", "#A8C7C7"), ("pressed", c["accent_hover"]), ("active", c["accent_hover"])],
            foreground=[("disabled", "#F5FAFA"), ("!disabled", "#FFFFFF")]
        )
        self.style.configure(
            "Danger.TButton", background="#FCEDEF", foreground=c["danger"],
            bordercolor="#F3D4D8", padding=(12, 9)
        )
        self.style.map("Danger.TButton", background=[("pressed", "#F4D9DD"), ("active", "#F8E4E7")])
        self.style.configure(
            "TEntry", fieldbackground=c["surface"], foreground=c["navy"],
            padding=(12, 11), bordercolor=c["border"], insertcolor=c["navy"],
            font=("Segoe UI", 11)
        )
        self.style.configure(
            "Treeview", rowheight=42, font=("Segoe UI", 10),
            background=c["surface"], fieldbackground=c["surface"],
            foreground=c["text"], bordercolor=c["border"], borderwidth=0
        )
        self.style.configure(
            "Treeview.Heading", font=("Segoe UI", 9, "bold"),
            background="#EDF2F8", foreground=c["navy"],
            padding=(12, 13), relief="flat"
        )
        self.style.map(
            "Treeview", background=[("selected", c["selected"])],
            foreground=[("selected", c["navy"])]
        )
        self.style.map("Treeview.Heading", background=[("active", "#DFE9F2")])
        self.style.configure(
            "Vertical.TScrollbar", background="#DDE6EF",
            troughcolor=c["background"], bordercolor=c["background"], arrowsize=15
        )

    def _build(self):
        outer = ttk.Frame(self.root, padding=(22, 18, 22, 14), style="App.TFrame")
        outer.pack(fill="both", expand=True)

        # Hero header
        hero = tk.Frame(outer, background=self.colors["hero"], height=142, bd=0, highlightthickness=0)
        hero.pack(fill="x", pady=(0, 15))
        hero.pack_propagate(False)
        hero_text = tk.Frame(hero, background=self.colors["hero"])
        hero_text.pack(side="right", fill="both", expand=True, padx=(24, 28), pady=19)
        tk.Label(
            hero_text, text="הספרייה האישית שלך",
            background=self.colors["hero"], foreground="#9EDBD7",
            font=("Segoe UI", 10, "bold"), anchor="e"
        ).pack(fill="x")
        tk.Label(
            hero_text, text=APP_NAME,
            background=self.colors["hero"], foreground="#FFFFFF",
            font=("Segoe UI", 27, "bold"), anchor="e"
        ).pack(fill="x", pady=(2, 1))
        tk.Label(
            hero_text, text="כל המתכונים, המסמכים והתמונות שלך — מסודרים ונגישים במקום אחד",
            background=self.colors["hero"], foreground="#C5D2E2",
            font=("Segoe UI", 10), anchor="e"
        ).pack(fill="x")

        offline_badge = tk.Frame(hero, background="#1D3A57", highlightbackground="#31536F", highlightthickness=1)
        offline_badge.pack(side="left", padx=25, pady=34)
        tk.Label(
            offline_badge, text="●  עובדת ללא אינטרנט",
            background="#1D3A57", foreground="#B8F0E7",
            font=("Segoe UI", 10, "bold"), padx=15, pady=12
        ).pack()
        tk.Frame(hero, background=self.colors["accent"], height=4, bd=0).pack(side="bottom", fill="x")

        # Search card
        search_frame = ttk.Frame(outer, padding=(14, 12), style="Card.TFrame")
        search_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(search_frame, text="חיפוש בספרייה", style="Section.TLabel").pack(side="right", padx=(10, 14))
        entry = ttk.Entry(search_frame, textvariable=self.search_value, justify="right")
        self.search_entry = entry
        entry.pack(side="right", fill="x", expand=True, padx=5)
        entry.bind("<KeyRelease>", lambda _event: self.refresh())
        ttk.Button(
            search_frame, text="ניקוי חיפוש",
            command=lambda: self.search_value.set("") or self.refresh()
        ).pack(side="left", padx=(5, 0))

        # Main actions
        toolbar = ttk.Frame(outer, style="App.TFrame")
        toolbar.pack(fill="x", pady=(0, 12))
        actions = [
            ("＋  הוספת קבצים", self.add_files),
            ("▣  ייבוא תיקייה", self.add_folder),
            ("📚  מרכז המתכונים", self.open_recipe_center),
            ("▤  פתיחת תיקיית הנתונים", self.open_data_folder),
            ("↥  יצירת גיבוי", self.backup),
            ("♻  סל מחזור", self.open_recycle_bin),
            ("⚙  הגדרות", self.settings),
        ]
        for label, command in actions:
            style = "Primary.TButton" if label.startswith("＋") else "TButton"
            ttk.Button(toolbar, text=label, command=command, style=style).pack(side="right", padx=(0, 7))

        # Quick filters for the entire library, favorites, and personal collections.
        filter_frame = ttk.Frame(outer, padding=(12, 9), style="Card.TFrame")
        filter_frame.pack(fill="x", pady=(0, 12))
        ttk.Label(filter_frame, text="הצגת קבצים:", style="Panel.TLabel").pack(side="right", padx=(0, 10))
        self.filter_combo = ttk.Combobox(
            filter_frame, textvariable=self.filter_value,
            state="readonly", width=30, justify="right"
        )
        self.filter_combo.pack(side="right", padx=4)
        self.filter_combo.bind("<<ComboboxSelected>>", self._on_filter_change)
        self.type_combo = ttk.Combobox(
            filter_frame, textvariable=self.type_filter_value,
            state="readonly", width=12, justify="right",
            values=["כל הסוגים", "תמונה", "מסמך", "גיליון", "מצגת", "ספר דיגיטלי", "קובץ"]
        )
        self.type_combo.pack(side="right", padx=4)
        self.type_combo.bind("<<ComboboxSelected>>", self._on_filter_change)
        self.category_combo = ttk.Combobox(
            filter_frame, textvariable=self.category_filter_value,
            state="readonly", width=15, justify="right"
        )
        self.category_combo.pack(side="right", padx=4)
        self.category_combo.bind("<<ComboboxSelected>>", self._on_filter_change)
        self.tag_combo = ttk.Combobox(
            filter_frame, textvariable=self.tag_filter_value,
            state="readonly", width=12, justify="right"
        )
        self.tag_combo.pack(side="right", padx=4)
        self.tag_combo.bind("<<ComboboxSelected>>", self._on_filter_change)
        self.sort_combo = ttk.Combobox(
            filter_frame, textvariable=self.sort_value,
            state="readonly", width=18, justify="right",
            values=["תאריך: חדש לישן", "תאריך: ישן לחדש", "שם: א-ת", "שם: ת-א", "סוג קובץ", "גודל: גדול לקטן", "דירוג: גבוה לנמוך", "נצפו לאחרונה"]
        )
        self.sort_combo.pack(side="right", padx=4)
        self.sort_combo.bind("<<ComboboxSelected>>", self._on_filter_change)
        ttk.Button(
            filter_frame, text="📁  ניהול אוספים",
            command=self.open_collections_dialog
        ).pack(side="left", padx=(4, 0))

        detail_filter_frame = ttk.Frame(outer, padding=(12, 7), style="Card.TFrame")
        detail_filter_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(detail_filter_frame, text="תיקייה:", style="Panel.TLabel").pack(side="right", padx=(0, 4))
        self.folder_combo = ttk.Combobox(detail_filter_frame, textvariable=self.folder_filter_value, state="readonly", width=25, justify="right", values=["כל התיקיות"])
        self.folder_combo.pack(side="right", padx=4)
        self.folder_combo.bind("<<ComboboxSelected>>", self._on_filter_change)
        ttk.Label(detail_filter_frame, text="סוג פריטים:", style="Panel.TLabel").pack(side="right", padx=(10, 4))
        self.recipe_filter_combo = ttk.Combobox(detail_filter_frame, textvariable=self.recipe_filter_value, state="readonly", width=21, justify="right", values=["כל הפריטים", "מתכונים בלבד", "מסמכים וקבצים", "מתכונים חסרים"])
        self.recipe_filter_combo.pack(side="right", padx=4)
        self.recipe_filter_combo.bind("<<ComboboxSelected>>", self._on_filter_change)
        ttk.Button(detail_filter_frame, text="בחירה מרובה / שינוי לקבוצה", command=self.manage_multiple_items).pack(side="left", padx=4)

        date_frame = ttk.Frame(outer, padding=(12, 7), style="Card.TFrame")
        date_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(date_frame, text="מתאריך:", style="Panel.TLabel").pack(side="right", padx=(0, 4))
        self.date_from_entry = ttk.Entry(date_frame, textvariable=self.date_from_value, width=11, justify="center")
        self.date_from_entry.pack(side="right", padx=3)
        ttk.Label(date_frame, text="עד תאריך:", style="Panel.TLabel").pack(side="right", padx=(6, 3))
        self.date_to_entry = ttk.Entry(date_frame, textvariable=self.date_to_value, width=11, justify="center")
        self.date_to_entry.pack(side="right", padx=3)
        self.date_from_entry.bind("<Return>", lambda _event: self.refresh())
        self.date_to_entry.bind("<Return>", lambda _event: self.refresh())
        ttk.Button(date_frame, text="החל תאריכים", command=self.refresh).pack(side="right", padx=4)
        ttk.Button(date_frame, text="ניקוי תאריכים", command=self._clear_date_filters).pack(side="right", padx=4)
        ttk.Button(date_frame, text="בדיקת כפילויות", command=self.show_duplicate_candidates).pack(side="left", padx=4)
        ttk.Button(date_frame, text="שחזור מגיבוי", command=self.restore_backup).pack(side="left", padx=4)

        # Resizable work area: preview on the left, file list on the right.
        body = ttk.Frame(outer, style="App.TFrame")
        body.pack(fill="both", expand=True)
        paned = tk.PanedWindow(
            body, orient="horizontal", background=self.colors["background"],
            sashwidth=8, sashrelief="flat", sashpad=2, bd=0, relief="flat"
        )
        paned.pack(fill="both", expand=True)

        preview = ttk.Frame(paned, padding=16, style="Card.TFrame")
        preview.configure(width=490)
        ttk.Label(preview, text="תצוגה מקדימה", style="Section.TLabel").pack(anchor="e", fill="x")
        ttk.Label(
            preview, text="בחרו תמונה מהרשימה כדי לראות אותה כאן בגודל גדול",
            style="Muted.TLabel"
        ).pack(anchor="e", fill="x", pady=(2, 10))

        self.preview_image = tk.Canvas(
            preview, background="#F2F6FA",
            highlightthickness=1, highlightbackground="#E3EAF2",
            bd=0, relief="flat"
        )
        self.preview_image.pack(fill="both", expand=True, pady=(0, 10))
        self.preview_image.bind("<Configure>", self._schedule_preview_resize)
        self._set_preview_placeholder("בחרו תמונה או קובץ מהרשימה")

        self.detail_label = ttk.Label(
            preview, text="", style="Muted.TLabel",
            justify="right", anchor="e", wraplength=430
        )
        self.detail_label.pack(fill="x", pady=(0, 9))
        selection_tools = ttk.Frame(preview, style="Card.TFrame")
        selection_tools.pack(fill="x", pady=(0, 7))
        self.favorite_button = ttk.Button(
            selection_tools, text="☆ הוספה למועדפים",
            command=self.toggle_favorite
        )
        self.favorite_button.pack(side="right", fill="x", expand=True, padx=(0, 4))
        ttk.Button(
            selection_tools, text="📁 שיוך לאוספים",
            command=self.open_collections_dialog
        ).pack(side="left", fill="x", expand=True, padx=(4, 0))
        ttk.Button(
            preview, text="תצוגת מתכון מקצועית  ✎",
            command=self.edit_recipe, style="Primary.TButton"
        ).pack(fill="x", pady=(0, 7))
        ttk.Button(
            preview, text="הדפסת מתכון / חוברת PDF  🖨",
            command=self.print_selected_recipe
        ).pack(fill="x", pady=(0, 7))
        ttk.Button(
            preview, text="פתיחת הקובץ  ↗", command=self.open_selected
        ).pack(fill="x", pady=(0, 7))
        preview_actions = ttk.Frame(preview, style="Card.TFrame")
        preview_actions.pack(fill="x")
        ttk.Button(preview_actions, text="שינוי שם", command=self.rename_selected).pack(side="right", fill="x", expand=True, padx=(0, 4))
        ttk.Button(
            preview_actions, text="הסרה מהספרייה",
            command=self.delete_selected, style="Danger.TButton"
        ).pack(side="left", fill="x", expand=True, padx=(4, 0))

        table_card = ttk.Frame(paned, padding=14, style="Card.TFrame")
        table_card.configure(width=790)
        table_header = ttk.Frame(table_card, style="Card.TFrame")
        table_header.pack(fill="x", pady=(0, 10))
        ttk.Label(table_header, text="הקבצים שלך", style="Section.TLabel").pack(side="right")
        ttk.Label(
            table_header, text="בחרו פריט כדי להציג פרטים ותצוגה מקדימה",
            style="Muted.TLabel"
        ).pack(side="left")
        view_switch = ttk.Frame(table_header, style="Card.TFrame")
        view_switch.pack(side="left", padx=(0, 12))
        ttk.Radiobutton(view_switch, text="☷ רשימה", value="רשימה", variable=self.view_mode, command=self.change_view_mode).pack(side="right", padx=3)
        ttk.Radiobutton(view_switch, text="▦ גלריה", value="גלריה", variable=self.view_mode, command=self.change_view_mode).pack(side="right", padx=3)

        table_body = ttk.Frame(table_card, style="Card.TFrame")
        table_body.pack(fill="both", expand=True)
        self.list_container = ttk.Frame(table_body, style="Card.TFrame")
        columns = ("favorite", "date", "kind", "size", "title")
        self.tree = ttk.Treeview(self.list_container, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("favorite", text="★", anchor="center")
        self.tree.heading("date", text="תאריך הוספה", anchor="e")
        self.tree.heading("kind", text="סוג קובץ", anchor="e")
        self.tree.heading("size", text="גודל", anchor="e")
        self.tree.heading("title", text="שם המתכון או הקובץ", anchor="e")
        self.tree.column("favorite", width=46, minwidth=42, anchor="center", stretch=False)
        self.tree.column("date", width=145, minwidth=125, anchor="e", stretch=False)
        self.tree.column("kind", width=115, minwidth=95, anchor="e", stretch=False)
        self.tree.column("size", width=95, minwidth=80, anchor="e", stretch=False)
        self.tree.column("title", width=470, minwidth=210, anchor="e", stretch=True)
        scroll = ttk.Scrollbar(self.list_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.tag_configure("even", background=self.colors["surface"])
        self.tree.tag_configure("odd", background=self.colors["row_alt"])
        self.tree.pack(side="right", fill="both", expand=True)
        scroll.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.show_selected)
        self.tree.bind("<Double-1>", lambda _event: self.open_selected())

        self.gallery_container = ttk.Frame(table_body, style="Card.TFrame")
        self.gallery_canvas = tk.Canvas(self.gallery_container, background=self.colors["surface"], highlightthickness=0, bd=0)
        gallery_scroll = ttk.Scrollbar(self.gallery_container, orient="vertical", command=self.gallery_canvas.yview)
        self.gallery_canvas.configure(yscrollcommand=gallery_scroll.set)
        gallery_scroll.pack(side="left", fill="y")
        self.gallery_canvas.pack(side="right", fill="both", expand=True)
        self.gallery_inner = ttk.Frame(self.gallery_canvas, style="Card.TFrame")
        self.gallery_window = self.gallery_canvas.create_window((0, 0), window=self.gallery_inner, anchor="nw")
        self.gallery_inner.bind("<Configure>", lambda _event: self.gallery_canvas.configure(scrollregion=self.gallery_canvas.bbox("all")))
        self.gallery_canvas.bind("<Configure>", lambda event: self.gallery_canvas.itemconfigure(self.gallery_window, width=event.width))
        if self.view_mode.get() == "גלריה":
            self.gallery_container.pack(fill="both", expand=True)
        else:
            self.list_container.pack(fill="both", expand=True)

        status_frame = ttk.Frame(outer, padding=(10, 7), style="Card.TFrame")
        status_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(status_frame, textvariable=self.status_value, style="Status.TLabel", anchor="e").pack(fill="x")

    def change_view_mode(self):
        """Switch between list and gallery views and remember the choice."""
        mode = self.view_mode.get()
        self.library.set_setting("view_mode", mode)
        if mode == "גלריה":
            self.list_container.pack_forget()
            self.gallery_container.pack(fill="both", expand=True)
        else:
            self.gallery_container.pack_forget()
            self.list_container.pack(fill="both", expand=True)
        self.refresh()

    def _bind_gallery_click(self, widget, item_id):
        widget.bind("<Button-1>", lambda _event, value=item_id: self._select_gallery_item(value))
        widget.bind("<Double-Button-1>", lambda _event, value=item_id: (self._select_gallery_item(value), self.open_selected()))
        widget.bind("<MouseWheel>", lambda event: self.gallery_canvas.yview_scroll(int(-event.delta / 120), "units"))
        for child in widget.winfo_children():
            self._bind_gallery_click(child, item_id)

    def _select_gallery_item(self, item_id):
        self.selected_item_id = str(item_id)
        self.show_selected()
        self._render_gallery_cards_selection()

    def _render_gallery_cards_selection(self):
        for item_id, card in getattr(self, "gallery_cards", {}).items():
            selected = str(item_id) == str(self.selected_item_id)
            try:
                card.configure(highlightbackground=self.colors["accent"] if selected else self.colors["border"], highlightthickness=3 if selected else 1)
            except tk.TclError:
                pass

    def _render_gallery(self, ranked):
        for child in self.gallery_inner.winfo_children():
            child.destroy()
        self.gallery_photos = []
        self.gallery_cards = {}
        columns = 4
        for index, (_score, row, file_size, _imported) in enumerate(ranked):
            item_id = str(row["id"])
            selected = item_id == str(self.selected_item_id)
            card = tk.Frame(self.gallery_inner, background=self.colors["surface"], highlightbackground=self.colors["accent"] if selected else self.colors["border"], highlightthickness=3 if selected else 1, bd=0, cursor="hand2")
            card.grid(row=index // columns, column=columns - 1 - (index % columns), sticky="nsew", padx=7, pady=7)
            self.gallery_cards[item_id] = card
            suffix = row["extension"].lower()
            preview = None
            if suffix in IMAGE_TYPES and Image is not None and ImageTk is not None:
                try:
                    with Image.open(row["stored_path"]) as source:
                        source.thumbnail((170, 112))
                        preview = ImageTk.PhotoImage(source.copy())
                    self.gallery_photos.append(preview)
                except Exception:
                    preview = None
            if preview is not None:
                picture = tk.Label(card, image=preview, background=self.colors["surface"], height=112)
                picture.image = preview
            else:
                picture = tk.Label(card, text=file_kind(suffix), background=self.colors["row_alt"], foreground=self.colors["accent"], font=("Segoe UI", 16, "bold"), height=4)
            picture.pack(fill="x", padx=8, pady=(8, 5))
            title = tk.Label(card, text=row["title"], background=self.colors["surface"], foreground=self.colors["navy"], font=("Segoe UI", 10, "bold"), justify="right", anchor="e", wraplength=170)
            title.pack(fill="x", padx=9, pady=(2, 4))
            size_label = f"{file_size / (1024 * 1024):.1f} MB" if file_size >= 1024 * 1024 else f"{max(1, file_size // 1024)} KB"
            details = tk.Label(card, text=f"{file_kind(suffix)}  ·  {row['imported_at']}  ·  {size_label}", background=self.colors["surface"], foreground=self.colors["muted"], font=("Segoe UI", 8), justify="right", anchor="e", wraplength=170)
            details.pack(fill="x", padx=9, pady=(0, 9))
            if row["is_favorite"]:
                favorite = tk.Label(card, text="★ מועדף", background=self.colors["surface"], foreground=self.colors["accent"], font=("Segoe UI", 8, "bold"))
                favorite.pack(anchor="e", padx=9, pady=(0, 5))
            self._bind_gallery_click(card, item_id)
        for column in range(columns):
            self.gallery_inner.grid_columnconfigure(column, weight=1, uniform="gallery")
        self.gallery_canvas.configure(scrollregion=self.gallery_canvas.bbox("all"))

    def _update_filter_options(self):
        options = {
            "כל הקבצים": ("all", None),
            "★ מועדפים": ("favorites", None),
        }
        for collection in self.library.collections():
            options[f"📁 {collection['name']}"] = ("collection", int(collection["id"]))
        current = self.filter_value.get()
        self._filter_options = options
        self.filter_combo.configure(values=list(options))
        self.filter_value.set(current if current in options else "כל הקבצים")
        categories = ["כל הקטגוריות"] + self.library.all_categories()
        tags = ["כל התגיות"] + self.library.all_tags()
        self.category_combo.configure(values=categories)
        self.tag_combo.configure(values=tags)
        if self.category_filter_value.get() not in categories:
            self.category_filter_value.set("כל הקטגוריות")
        if self.tag_filter_value.get() not in tags:
            self.tag_filter_value.set("כל התגיות")
        folders = ["כל התיקיות"] + self.library.all_source_folders()
        if hasattr(self, "folder_combo"):
            self.folder_combo.configure(values=folders)
        if self.folder_filter_value.get() not in folders:
            self.folder_filter_value.set("כל התיקיות")

    def _on_filter_change(self, _event=None):
        self.refresh()

    def _clear_date_filters(self):
        self.date_from_value.set("")
        self.date_to_value.set("")
        self.refresh()

    @staticmethod
    def _parse_imported_date(value: str):
        try:
            return datetime.strptime((value or "").strip(), "%d/%m/%Y %H:%M")
        except ValueError:
            return None

    @staticmethod
    def _parse_date_filter(value: str):
        value = (value or "").strip()
        if not value:
            return None
        try:
            return datetime.strptime(value, "%d/%m/%Y").date()
        except ValueError:
            return None

    def refresh(self):
        selected = self.tree.selection()
        selected_id = selected[0] if selected else self.selected_item_id
        self._update_filter_options()
        mode, filter_id = self._filter_options.get(self.filter_value.get(), ("all", None))
        rows = self.library.rows(
            favorite_only=(mode == "favorites"),
            collection_id=(filter_id if mode == "collection" else None)
        )
        query = self.search_value.get().strip()
        type_filter = self.type_filter_value.get()
        category_filter = self.category_filter_value.get()
        tag_filter = self.tag_filter_value.get()
        date_from = self._parse_date_filter(self.date_from_value.get())
        date_to = self._parse_date_filter(self.date_to_value.get())
        ranked = []
        for row in rows:
            kind = file_kind(row["extension"])
            if type_filter != "כל הסוגים" and kind != type_filter:
                continue
            recipe = self.library.get_recipe_data(row["id"])
            category = recipe.get("category", "")
            folder_path = str(recipe.get("source_folder") or "").strip("/")
            folder_filter = self.folder_filter_value.get()
            if folder_filter != "כל התיקיות" and not (folder_path == folder_filter or folder_path.startswith(folder_filter + "/")):
                continue
            recipe_filter = self.recipe_filter_value.get()
            is_recipe = is_recipe_record(recipe, row["content"] or "")
            if recipe_filter == "מתכונים בלבד" and not is_recipe:
                continue
            if recipe_filter == "מסמכים וקבצים" and is_recipe:
                continue
            if recipe_filter == "מתכונים חסרים" and (not is_recipe or not recipe_missing_parts(recipe)):
                continue
            if category_filter != "כל הקטגוריות" and category != category_filter:
                continue
            if tag_filter != "כל התגיות" and tag_filter not in self.library.tag_names(row["id"]):
                continue
            imported = self._parse_imported_date(row["imported_at"])
            if date_from and (imported is None or imported.date() < date_from):
                continue
            if date_to and (imported is None or imported.date() > date_to):
                continue
            item_tags = self.library.tag_names(row["id"])
            search_text = recipe_search_text(
                row["title"], row["content"], recipe, item_tags
            )
            score = relevance(row["title"], search_text, query)
            if score > 0:
                try:
                    file_size = Path(row["stored_path"]).stat().st_size
                except OSError:
                    file_size = 0
                ranked.append((score, row, file_size, imported))
        sort_mode = self.sort_value.get()
        if sort_mode == "שם: א-ת":
            ranked.sort(key=lambda x: x[1]["title"].casefold())
        elif sort_mode == "שם: ת-א":
            ranked.sort(key=lambda x: x[1]["title"].casefold(), reverse=True)
        elif sort_mode == "תאריך: ישן לחדש":
            ranked.sort(key=lambda x: x[3] or datetime.min)
        elif sort_mode == "סוג קובץ":
            ranked.sort(key=lambda x: (file_kind(x[1]["extension"]), x[1]["title"].casefold()))
        elif sort_mode == "גודל: גדול לקטן":
            ranked.sort(key=lambda x: x[2], reverse=True)
        elif sort_mode == "דירוג: גבוה לנמוך":
            ranked.sort(
                key=lambda x: int(self.library.get_recipe_data(x[1]["id"]).get("rating") or 0)
                if str(self.library.get_recipe_data(x[1]["id"]).get("rating") or "0").isdigit() else 0,
                reverse=True
            )
        elif sort_mode == "נצפו לאחרונה":
            ranked.sort(key=lambda x: x[1]["last_opened"] or "", reverse=True)
        elif query:
            ranked.sort(key=lambda x: (x[0], x[1]["id"]), reverse=True)
        else:
            ranked.sort(key=lambda x: (x[3] or datetime.min, x[1]["id"]), reverse=True)

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.current_rows = {}
        for index, (_score, row, file_size, _imported) in enumerate(ranked):
            iid = str(row["id"])
            self.current_rows[iid] = row
            star = "★" if row["is_favorite"] else ""
            if file_size >= 1024 * 1024:
                size_label = f"{file_size / (1024 * 1024):.1f} MB"
            else:
                size_label = f"{max(1, file_size // 1024)} KB"
            self.tree.insert(
                "", "end", iid=iid,
                values=(star, row["imported_at"], file_kind(row["extension"]), size_label, row["title"]),
                tags=("even" if index % 2 == 0 else "odd",)
            )
        self._render_gallery(ranked)
        self.status_value.set(f"נמצאו {len(ranked)} פריטים לפי הסינון הנוכחי.")
        if selected_id and selected_id in self.current_rows:
            self.selected_item_id = selected_id
            if self.view_mode.get() == "רשימה":
                self.tree.selection_set(selected_id)
                self.tree.focus(selected_id)
            self.show_selected()
            self._render_gallery_cards_selection()
        else:
            self.selected_item_id = None
            self.preview_path = None
            self.preview_render_path = None
            self.preview_photo = None
            self._cleanup_preview_temp()
            if self._preview_resize_job is not None:
                try:
                    self.root.after_cancel(self._preview_resize_job)
                except tk.TclError:
                    pass
                self._preview_resize_job = None
            self._set_preview_placeholder("בחרו תמונה או קובץ מהרשימה")
            self.detail_label.configure(text="")
            self._update_selection_actions(None)

    def selected_row(self):
        if self.view_mode.get() == "גלריה":
            return self.current_rows.get(str(self.selected_item_id)) if self.selected_item_id else None
        selected = self.tree.selection()
        if selected:
            self.selected_item_id = selected[0]
            return self.current_rows.get(selected[0])
        return None

    def _set_preview_placeholder(self, message: str):
        """Draw placeholder text on the canvas without affecting its requested size."""
        if not hasattr(self, "preview_image"):
            return
        self.preview_image.delete("all")
        width = max(1, self.preview_image.winfo_width())
        height = max(1, self.preview_image.winfo_height())
        self.preview_image.create_text(
            width // 2, height // 2,
            text=message, fill=self.colors["muted"],
            font=("Segoe UI", 12), justify="center",
            width=max(160, width - 40), tags=("preview-placeholder",)
        )

    def _update_selection_actions(self, row):
        if self.favorite_button is None:
            return
        if row is None:
            self.favorite_button.configure(text="☆ הוספה למועדפים", state="disabled")
        elif row["is_favorite"]:
            self.favorite_button.configure(text="★ הסרה מהמועדפים", state="normal")
        else:
            self.favorite_button.configure(text="☆ הוספה למועדפים", state="normal")

    def show_selected(self, _event=None):
        row = self.selected_row()
        if not row:
            self._update_selection_actions(None)
            return
        self.selected_item_id = str(row["id"])
        path = Path(row["stored_path"])
        collection_names = self.library.item_collection_names(row["id"])
        collection_text = ", ".join(collection_names) if collection_names else "ללא אוסף"
        tag_names = self.library.tag_names(row["id"])
        category = self.library.get_recipe_data(row["id"]).get("category", "")
        info = (
            f"שם: {row['title']}\n"
            f"קובץ מקורי: {row['original_name']}\n"
            f"נוסף בתאריך: {row['imported_at']}\n"
            f"קטגוריה: {category or 'ללא קטגוריה'}\n"
            f"תיקייה: {self.library.get_recipe_data(row['id']).get('source_folder') or 'שורש הספרייה'}\n"
            f"דירוג אישי: {self.library.get_recipe_data(row['id']).get('rating') or 'ללא דירוג'}\n"
            f"תגיות: {', '.join(tag_names) if tag_names else 'ללא תגיות'}\n"
            f"אוספים: {collection_text}"
        )
        self.detail_label.configure(text=info)
        self._update_selection_actions(row)
        self._cleanup_preview_temp()
        self.preview_photo = None
        self.preview_path = path
        self.preview_render_path = path
        self.preview_notice = ""
        if path.suffix.lower() in OFFICE_TYPES:
            converted, temp_dir, notice = convert_office_to_pdf(path)
            if converted is not None:
                self.preview_render_path = converted
                self.preview_temp_directory = temp_dir
            else:
                self.preview_notice = notice
        self._set_preview_placeholder("טוען תצוגה מקדימה…")
        self.root.after_idle(self._render_selected_preview)

    def toggle_favorite(self):
        row = self.selected_row()
        if not row:
            return
        self.library.set_favorite(row["id"], not bool(row["is_favorite"]))
        self.status_value.set(
            "הפריט נוסף למועדפים." if not row["is_favorite"] else "הפריט הוסר מהמועדפים."
        )
        self.refresh()

    def open_collections_dialog(self):
        row = self.selected_row()
        item_id = int(row["id"]) if row else None
        dialog = tk.Toplevel(self.root)
        dialog.title("ניהול אוספים")
        dialog.geometry("500x540")
        dialog.minsize(420, 360)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(background=self.colors["background"])

        outer = ttk.Frame(dialog, padding=18, style="App.TFrame")
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="האוספים שלי", style="Section.TLabel").pack(fill="x", anchor="e")
        description = (
            f"סמנו באילו אוספים יופיע הפריט: {row['title']}"
            if row else "כאן אפשר ליצור אוספים, למחוק אוספים ולנהל את החלוקה האישית."
        )
        ttk.Label(
            outer, text=description, style="Muted.TLabel",
            wraplength=440, justify="right"
        ).pack(fill="x", anchor="e", pady=(4, 12))

        list_frame = ttk.Frame(outer, padding=10, style="Card.TFrame")
        list_frame.pack(fill="both", expand=True)

        def rebuild_list():
            for child in list_frame.winfo_children():
                child.destroy()
            collections = self.library.collections()
            if not collections:
                ttk.Label(
                    list_frame, text="עדיין לא יצרתם אוספים.\nלחצו על ״יצירת אוסף חדש״ כדי להתחיל.",
                    style="Muted.TLabel", justify="center"
                ).pack(fill="both", expand=True, pady=22)
                return
            assigned = self.library.item_collection_ids(item_id) if item_id is not None else set()
            for collection in collections:
                collection_id = int(collection["id"])
                line = ttk.Frame(list_frame, style="Card.TFrame")
                line.pack(fill="x", pady=4)
                if item_id is not None:
                    variable = tk.BooleanVar(value=collection_id in assigned)
                    def on_toggle(cid=collection_id, var=variable):
                        self.library.set_item_collection(item_id, cid, var.get())
                        self.refresh()
                    ttk.Checkbutton(
                        line, text=collection["name"], variable=variable,
                        command=on_toggle
                    ).pack(side="right", fill="x", expand=True)
                else:
                    ttk.Label(
                        line, text=collection["name"], style="Panel.TLabel"
                    ).pack(side="right", fill="x", expand=True)
                def delete_collection(cid=collection_id, name=collection["name"]):
                    if messagebox.askyesno(
                        "מחיקת אוסף",
                        f"למחוק את האוסף ״{name}״?\nהקבצים עצמם לא יימחקו.",
                        parent=dialog
                    ):
                        self.library.delete_collection(cid)
                        self._update_filter_options()
                        self.refresh()
                        rebuild_list()
                ttk.Button(
                    line, text="מחיקה", style="Danger.TButton",
                    command=delete_collection
                ).pack(side="left", padx=(6, 0))

        def create_collection():
            name = simpledialog.askstring(
                "יצירת אוסף", "מה יהיה שם האוסף החדש?",
                parent=dialog
            )
            if name is None:
                return
            try:
                self.library.create_collection(name)
            except ValueError as exc:
                messagebox.showwarning("לא ניתן ליצור אוסף", str(exc), parent=dialog)
                return
            self._update_filter_options()
            rebuild_list()
            self.status_value.set("האוסף החדש נוצר.")

        rebuild_list()
        buttons = ttk.Frame(outer, style="App.TFrame")
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(
            buttons, text="＋ יצירת אוסף חדש",
            style="Primary.TButton", command=create_collection
        ).pack(side="right")
        ttk.Button(buttons, text="סגירה", command=dialog.destroy).pack(side="left")

    def _schedule_preview_resize(self, _event=None):
        if self.preview_path is None:
            return
        if self._preview_resize_job is not None:
            try:
                self.root.after_cancel(self._preview_resize_job)
            except tk.TclError:
                pass
        self._preview_resize_job = self.root.after(160, self._render_selected_preview)

    def _render_selected_preview(self):
        self._preview_resize_job = None
        path = self.preview_render_path or self.preview_path
        source_path = self.preview_path
        if path is None or not path.is_file():
            self._set_preview_placeholder("הקובץ לא נמצא.")
            return
        try:
            self.preview_image.update_idletasks()
            width = self.preview_image.winfo_width()
            height = self.preview_image.winfo_height()
            if width < 80 or height < 80:
                return
            target_width, target_height = max(1, width - 28), max(1, height - 28)
            suffix = source_path.suffix.lower()

            if suffix in IMAGE_TYPES and Image is not None and ImageTk is not None:
                with Image.open(source_path) as source:
                    source.seek(0)
                    image = source.convert("RGBA")
                if fitz is not None:
                    zoom = float(self.library.get_setting("preview_zoom", "1.0") or "1.0")
                    target_width = max(1, int(target_width * zoom))
                    target_height = max(1, int(target_height * zoom))
                resampling = getattr(Image, "Resampling", Image)
                image.thumbnail((target_width, target_height), getattr(resampling, "LANCZOS", 1))
                self._put_preview_image(ImageTk.PhotoImage(image), width, height)
                return

            is_pdf_preview = (
                path.suffix.lower() == ".pdf"
                or suffix in EBOOK_TYPES
                or (suffix in OFFICE_TYPES and path != source_path)
            )
            if is_pdf_preview and fitz is not None and Image is not None and ImageTk is not None:
                with fitz.open(path) as document:
                    if document.page_count:
                        page = document[0]
                        zoom = max(0.5, min(1.6, float(self.library.get_setting("preview_zoom", "1.0") or "1.0")))
                        base_scale = min(target_width / max(1, page.rect.width), target_height / max(1, page.rect.height))
                        scale = max(0.1, min(2.5, base_scale * zoom))
                        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                        image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                        self._put_preview_image(ImageTk.PhotoImage(image), width, height)
                        return

            content = ""
            try:
                row = self.selected_row()
                content = (row["content"] if row else "") or extract_text(source_path)
            except Exception:
                content = ""
            if content.strip():
                self.preview_image.delete("all")
                if self.preview_notice:
                    content = self.preview_notice + "\n\n" + content
                self.preview_image.create_text(
                    width - 20, 18, anchor="ne",
                    text=content[:1100], fill=self.colors["text"],
                    font=("Segoe UI", 10), justify="right",
                    width=max(160, width - 40), tags=("preview-text",)
                )
                return

            description = file_kind(suffix) + "\n" + (source_path.name[:80])
            if self.preview_notice:
                description += "\n\n" + self.preview_notice
            self._set_preview_placeholder(description)
        except Exception as exc:
            self._set_preview_placeholder(f"לא ניתן להציג תצוגה מקדימה.\n{exc}")

    def _put_preview_image(self, photo, width: int, height: int):
        self.preview_photo = photo
        self.preview_image.delete("all")
        self.preview_image.create_image(
            width // 2, height // 2,
            image=self.preview_photo, anchor="center",
            tags=("preview-image",)
        )

    def _bind_shortcuts(self):
        self.root.bind_all("<Control-f>", self._focus_search)
        self.root.bind_all("<Control-F>", self._focus_search)
        self.root.bind_all("<Control-o>", lambda _event: self.open_selected())
        self.root.bind_all("<Control-O>", lambda _event: self.open_selected())
        self.root.bind_all("<Control-b>", lambda _event: self.backup())
        self.root.bind_all("<Control-B>", lambda _event: self.backup())
        self.root.bind_all("<Control-Shift-n>", lambda _event: self.open_collections_dialog())
        self.root.bind_all("<Escape>", self._clear_search_shortcut)

    def _focus_search(self, _event=None):
        if self.search_entry is not None:
            self.search_entry.focus_set()
            self.search_entry.selection_range(0, "end")
        return "break"

    def _clear_search_shortcut(self, _event=None):
        widget = self.root.focus_get()
        if widget is self.search_entry and self.search_value.get():
            self.search_value.set("")
            self.refresh()
            return "break"
        return None

    def _maybe_auto_backup(self):
        try:
            enabled = self.library.get_setting("auto_backup_enabled", "0") == "1"
            interval = max(1, min(90, int(self.library.get_setting("auto_backup_days", "7"))))
            last_value = self.library.get_setting("last_auto_backup", "")
            last_time = None
            if last_value:
                try:
                    last_time = datetime.fromisoformat(last_value)
                except ValueError:
                    last_time = None
            due = last_time is None or datetime.now() - last_time >= timedelta(days=interval)
            if enabled and due:
                target = self.library.backup_dir / (
                    "גיבוי-אוטומטי-" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".zip"
                )
                self.library.create_backup(target)
                self.library.set_setting("last_auto_backup", datetime.now().isoformat(timespec="seconds"))
                backups = sorted(
                    self.library.backup_dir.glob("גיבוי-אוטומטי-*.zip"),
                    key=lambda p: p.stat().st_mtime,
                    reverse=True
                )
                for old in backups[5:]:
                    old.unlink(missing_ok=True)
                self.status_value.set("נוצר גיבוי אוטומטי של הספרייה.")
        except Exception as exc:
            self.status_value.set(f"הגיבוי האוטומטי לא הושלם: {exc}")
        finally:
            self._auto_backup_job = self.root.after(6 * 60 * 60 * 1000, self._maybe_auto_backup)

    def manage_multiple_items(self):
        if self.view_mode.get() != "רשימה":
            messagebox.showinfo("ניהול מרובה", "עברו לתצוגת רשימה כדי לבחור כמה פריטים בעזרת Ctrl או Shift.")
            return
        selected = list(self.tree.selection())
        if not selected:
            messagebox.showinfo("ניהול מרובה", "סמנו כמה מתכונים או קבצים ברשימה, ואז לחצו שוב.")
            return
        category = simpledialog.askstring(
            "שינוי קטגוריה לקבוצה",
            f"נבחרו {len(selected)} פריטים. הזינו קטגוריה חדשה. השארת השדה ריק תשאיר את הקטגוריות הקיימות.",
            parent=self.root
        )
        if category is None:
            return
        tags = simpledialog.askstring(
            "שינוי תגיות לקבוצה",
            "הזינו תגיות מופרדות בפסיקים. השארת השדה ריק תשאיר את התגיות הקיימות.",
            parent=self.root
        )
        if tags is None:
            return
        category = category.strip()
        tag_names = [part.strip() for part in tags.split(",") if part.strip()]
        if not category and not tag_names:
            messagebox.showinfo("ניהול מרובה", "לא הוזנו שינויים.")
            return
        changed = 0
        for item in selected:
            try:
                item_id = int(item)
                data = self.library.get_recipe_data(item_id)
                if category:
                    data["category"] = category
                self.library.set_recipe_data(item_id, data)
                if tag_names:
                    self.library.set_item_tags(item_id, tag_names)
                changed += 1
            except Exception:
                continue
        self.refresh()
        self.status_value.set(f"העדכון הקבוצתי הושלם עבור {changed} פריטים.")

    def edit_recipe(self):
        row = self.selected_row()
        if not row:
            messagebox.showinfo("פרטי מתכון", "יש לבחור קובץ או מתכון מהרשימה.")
            return
        item_id = int(row["id"])
        data = self.library.get_recipe_data(item_id)
        tags = self.library.tag_names(item_id)

        window = tk.Toplevel(self.root)
        window.title(f"פרטי מתכון – {row['title']}")
        window.geometry("900x820")
        window.minsize(760, 650)
        window.transient(self.root)
        window.configure(background=self.colors["background"])

        outer = ttk.Frame(window, padding=18, style="App.TFrame")
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="פרטי המתכון", style="Section.TLabel").pack(fill="x", anchor="e")
        title_frame = ttk.Frame(outer, padding=10, style="Card.TFrame")
        title_frame.pack(fill="x", pady=(6, 0))
        ttk.Label(title_frame, text="שם המתכון", style="Muted.TLabel").pack(anchor="e")
        title_var = tk.StringVar(value=row["title"])
        ttk.Entry(title_frame, textvariable=title_var, justify="right").pack(fill="x", pady=(3, 0))
        ttk.Label(
            outer,
            text="שמרו את פרטי המתכון, המצרכים, הוראות ההכנה והתגיות האישיות.",
            style="Muted.TLabel"
        ).pack(fill="x", pady=(3, 8))

        top = ttk.Frame(outer, padding=12, style="Card.TFrame")
        top.pack(fill="x")
        variables = {}
        fields = [
            ("category", "קטגוריה", data.get("category", ""), 22),
            ("prep_minutes", "זמן הכנה בדקות", str(data.get("prep_minutes", "")), 10),
            ("cook_minutes", "זמן בישול בדקות", str(data.get("cook_minutes", "")), 10),
            ("servings", "מספר מנות", str(data.get("servings", "4") or "4"), 8),
            ("rating", "דירוג אישי (0–5)", str(data.get("rating", "") or ""), 8),
        ]
        for key, label, value, width in fields:
            cell = ttk.Frame(top, style="Card.TFrame")
            cell.pack(side="right", padx=5)
            ttk.Label(cell, text=label, style="Muted.TLabel").pack(anchor="e")
            variable = tk.StringVar(value=value)
            variables[key] = variable
            ttk.Entry(cell, textvariable=variable, width=width, justify="right").pack(anchor="e", pady=(3, 0))

        tags_frame = ttk.Frame(outer, padding=12, style="Card.TFrame")
        tags_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(tags_frame, text="תגיות אישיות (הפרידו בפסיקים)", style="Muted.TLabel").pack(anchor="e")
        tags_var = tk.StringVar(value=", ".join(tags))
        ttk.Entry(tags_frame, textvariable=tags_var, justify="right").pack(fill="x", pady=(4, 0))
        recipe_flag = tk.BooleanVar(value=is_recipe_record(data, row["content"] or ""))
        ttk.Checkbutton(tags_frame, text="זהו מתכון (ולא מסמך או קובץ כללי)", variable=recipe_flag).pack(anchor="e", pady=(7, 0))
        ttk.Label(tags_frame, text=f"תיקיית מקור: {data.get('source_folder') or 'שורש הספרייה'}", style="Muted.TLabel").pack(anchor="e", pady=(4, 0))
        image_path_var = tk.StringVar(value=str(data.get("image_path") or (row["stored_path"] if row["extension"] in IMAGE_TYPES else "")))
        image_source_var = tk.StringVar(value="")
        image_frame = ttk.Frame(outer, padding=9, style="Card.TFrame")
        image_frame.pack(fill="x", pady=(8, 0))
        ttk.Label(image_frame, text="תמונה מייצגת למתכון", style="Panel.TLabel").pack(side="right", padx=(0, 8))
        image_label = ttk.Label(image_frame, text=Path(image_path_var.get()).name if image_path_var.get() else "לא נבחרה תמונה", style="Muted.TLabel")
        image_label.pack(side="right", padx=8)
        def choose_recipe_image():
            selected_image = filedialog.askopenfilename(
                parent=window, title="בחירת תמונה למתכון",
                filetypes=[("תמונות", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp")]
            )
            if selected_image:
                image_source_var.set(selected_image)
                image_label.configure(text=Path(selected_image).name)
        ttk.Button(image_frame, text="בחירת תמונה…", command=choose_recipe_image).pack(side="left")

        def add_text_section(parent, title, initial, height):
            frame = ttk.Frame(parent, padding=10, style="Card.TFrame")
            frame.pack(fill="both", expand=True, pady=(8, 0))
            ttk.Label(frame, text=title, style="Section.TLabel").pack(fill="x", anchor="e", pady=(0, 5))
            area = tk.Text(
                frame, height=height, wrap="word", font=("Segoe UI", 10),
                background="#FFFFFF", foreground=self.colors["text"],
                insertbackground=self.colors["accent"], relief="flat",
                highlightthickness=1, highlightbackground=self.colors["border"],
                padx=10, pady=8
            )
            area.pack(fill="both", expand=True)
            area.insert("1.0", initial or "")
            area.tag_configure("rtl", justify="right")
            area.tag_add("rtl", "1.0", "end")
            return frame, area

        content = row["content"] or ""
        initial_ingredients = data.get("ingredients") or extract_ingredients(content)
        ingredients_frame, ingredients_text = add_text_section(
            outer, "מצרכים – שורה לכל רכיב", initial_ingredients, 8
        )
        tools = ttk.Frame(ingredients_frame, style="Card.TFrame")
        tools.pack(fill="x", pady=(6, 0))
        ttk.Button(
            tools, text="זיהוי מצרכים מתוך הקובץ",
            command=lambda: (
                ingredients_text.delete("1.0", "end"),
                ingredients_text.insert("1.0", extract_ingredients(content))
            )
        ).pack(side="right")

        def smart_import_recipe():
            parsed = parse_recipe_text(content)
            if parsed.get("title"):
                title_var.set(parsed["title"])
            if parsed.get("ingredients"):
                ingredients_text.delete("1.0", "end")
                ingredients_text.insert("1.0", parsed["ingredients"])
            if parsed.get("instructions"):
                instructions_text.delete("1.0", "end")
                instructions_text.insert("1.0", parsed["instructions"])
            messagebox.showinfo(
                "ייבוא חכם",
                "הטקסט נותח. בדקו את המצרכים וההוראות לפני השמירה."
                if parsed.get("is_recipe") else
                "לא זוהו בביטחון כותרות של מצרכים והוראות. הטקסט המקורי נשאר שמור.",
                parent=window
            )
        ttk.Button(tools, text="ייבוא חכם של המתכון", command=smart_import_recipe).pack(side="left")

        instructions_frame, instructions_text = add_text_section(
            outer, "הוראות הכנה", data.get("instructions", ""), 8
        )
        notes_frame, notes_text = add_text_section(
            outer, "הערות אישיות", data.get("notes", ""), 4
        )

        scale_frame = ttk.Frame(outer, padding=10, style="Card.TFrame")
        scale_frame.pack(fill="x", pady=(9, 0))
        ttk.Label(scale_frame, text="חישוב כמויות למספר מנות אחר", style="Panel.TLabel").pack(side="right", padx=(0, 8))
        target_servings = tk.StringVar(value=str(data.get("servings", "4") or "4"))
        ttk.Entry(scale_frame, textvariable=target_servings, width=7, justify="center").pack(side="right", padx=5)

        def show_scaled_ingredients():
            try:
                base = int(variables["servings"].get() or "4")
                target = int(target_servings.get())
                if base < 1 or target < 1 or target > 500:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("כמויות", "הזינו מספר מנות תקין (1 עד 500).", parent=window)
                return
            current = ingredients_text.get("1.0", "end-1c")
            scaled = scale_ingredient_amounts(current, base, target)
            result = tk.Toplevel(window)
            result.title(f"כמויות ל־{target} מנות")
            result.geometry("520x560")
            frame = ttk.Frame(result, padding=14, style="App.TFrame")
            frame.pack(fill="both", expand=True)
            ttk.Label(frame, text=f"המצרכים ל־{target} מנות", style="Section.TLabel").pack(fill="x")
            text_area = tk.Text(frame, wrap="word", font=("Segoe UI", 11), padx=12, pady=12)
            text_area.pack(fill="both", expand=True, pady=8)
            text_area.insert("1.0", scaled)
            text_area.tag_configure("rtl", justify="right")
            text_area.tag_add("rtl", "1.0", "end")
            text_area.configure(state="disabled")

        ttk.Button(scale_frame, text="חשב כמויות", command=show_scaled_ingredients).pack(side="left")

        buttons = ttk.Frame(outer, style="App.TFrame")
        buttons.pack(fill="x", pady=(12, 0))

        def save_recipe_version():
            label = simpledialog.askstring("שמירת גרסה", "שם לגרסה החדשה (למשל: כמות כפולה או פחות סוכר):", parent=window)
            if not label or not label.strip():
                return
            current = self.library.get_recipe_data(item_id)
            versions = list(current.get("versions") or [])
            versions.append({
                "label": label.strip(),
                "saved_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "ingredients": ingredients_text.get("1.0", "end-1c").strip(),
                "instructions": instructions_text.get("1.0", "end-1c").strip(),
                "servings": variables["servings"].get().strip(),
            })
            current["versions"] = versions[-30:]
            self.library.set_recipe_data(item_id, current)
            messagebox.showinfo("שמירת גרסה", "הגרסה נשמרה בנפרד. המתכון הראשי לא הוחלף.", parent=window)

        def compare_recipe_versions():
            versions = list(self.library.get_recipe_data(item_id).get("versions") or [])
            if not versions:
                messagebox.showinfo("השוואת גרסאות", "עדיין לא נשמרו גרסאות למתכון.", parent=window)
                return
            labels = [f"{index + 1}. {item.get('label','גרסה')} · {item.get('saved_at','')}" for index, item in enumerate(versions)]
            choice = simpledialog.askinteger("השוואת גרסאות", "בחרו את מספר הגרסה להשוואה:\n" + "\n".join(labels), parent=window, minvalue=1, maxvalue=len(versions))
            if not choice:
                return
            version = versions[choice - 1]
            comparison = (
                f"גרסה שמורה: {version.get('label','')} · {version.get('saved_at','')}\n\n"
                "מצרכים – הגרסה השמורה:\n" + str(version.get("ingredients") or "") + "\n\n"
                "מצרכים – הגרסה הנוכחית בחלון:\n" + ingredients_text.get("1.0", "end-1c").strip() + "\n\n"
                "הוראות – הגרסה השמורה:\n" + str(version.get("instructions") or "") + "\n\n"
                "הוראות – הגרסה הנוכחית בחלון:\n" + instructions_text.get("1.0", "end-1c").strip()
            )
            result = tk.Toplevel(window)
            result.title("השוואת גרסאות")
            result.geometry("760x650")
            text_area = tk.Text(result, wrap="word", font=("Segoe UI", 10), padx=12, pady=12)
            text_area.pack(fill="both", expand=True)
            text_area.insert("1.0", comparison)
            text_area.tag_configure("rtl", justify="right")
            text_area.tag_add("rtl", "1.0", "end")
            text_area.configure(state="disabled")

        ttk.Button(buttons, text="שמירת גרסה נפרדת", command=save_recipe_version).pack(side="right", padx=3)
        ttk.Button(buttons, text="השוואת גרסאות", command=compare_recipe_versions).pack(side="right", padx=3)

        def save_recipe():
            try:
                for key in ("prep_minutes", "cook_minutes", "servings"):
                    value = variables[key].get().strip()
                    if value and (not value.isdigit() or int(value) > 10000):
                        raise ValueError("זמני הכנה, בישול ומספר מנות חייבים להיות מספרים חיוביים.")
                if variables["servings"].get().strip() and int(variables["servings"].get()) < 1:
                    raise ValueError("מספר המנות חייב להיות לפחות 1.")
                rating = variables["rating"].get().strip()
                if rating and (not rating.isdigit() or not 0 <= int(rating) <= 5):
                    raise ValueError("הדירוג האישי חייב להיות מספר בין 0 ל־5.")
                recipe = {
                    "category": variables["category"].get().strip(),
                    "prep_minutes": variables["prep_minutes"].get().strip(),
                    "cook_minutes": variables["cook_minutes"].get().strip(),
                    "servings": variables["servings"].get().strip() or "4",
                    "ingredients": ingredients_text.get("1.0", "end-1c").strip(),
                    "instructions": instructions_text.get("1.0", "end-1c").strip(),
                    "notes": notes_text.get("1.0", "end-1c").strip(),
                    "rating": rating,
                    "is_recipe": bool(recipe_flag.get()),
                    "image_path": image_path_var.get().strip(),
                }
                if image_source_var.get().strip():
                    selected_image = Path(image_source_var.get())
                    try:
                        selected_image.resolve().relative_to(self.library.files_dir.resolve())
                        stored_image = selected_image
                    except ValueError:
                        image_folder = self.library.files_dir / "תמונות מתכונים"
                        image_folder.mkdir(parents=True, exist_ok=True)
                        stored_image = unique_path(image_folder, f"{item_id}-{safe_title(row['title'])}", selected_image.suffix.lower())
                        shutil.copy2(selected_image, stored_image)
                    recipe["image_path"] = str(stored_image)
                new_title = safe_title(title_var.get().strip())
                if new_title and new_title != row["title"]:
                    self.library.rename(item_id, new_title)
                self.library.set_recipe_data(item_id, recipe)
                self.library.set_item_tags(item_id, tags_var.get().split(","))
                self.refresh()
                self.status_value.set("פרטי המתכון והתגיות נשמרו.")
                window.destroy()
            except Exception as exc:
                messagebox.showerror("שמירת מתכון", str(exc), parent=window)

        ttk.Button(buttons, text="שמירת הפרטים", style="Primary.TButton", command=save_recipe).pack(side="right")
        ttk.Button(buttons, text="סגירה", command=window.destroy).pack(side="left")


    def open_recipe_center(self, initial_item_id=None):
        RecipeCenter(
            self, extract_ingredients, file_kind, safe_title, convert_office_to_pdf
        ).show(initial_item_id=initial_item_id)

    def print_selected_recipe(self):
        row = self.selected_row()
        if not row:
            messagebox.showinfo("הדפסת מתכון", "יש לבחור מתכון או פריט בספרייה.")
            return
        self.open_recipe_center(initial_item_id=int(row["id"]))

    def show_duplicate_candidates(self):
        rows = self.library.rows()
        if len(rows) < 2:
            messagebox.showinfo("בדיקת כפילויות", "נדרשים לפחות שני קבצים בספרייה.")
            return
        # Avoid quadratic work on extremely large libraries; compare the newest 700 items.
        rows = rows[:350]
        found = []
        for index, first in enumerate(rows):
            first_title = normalize(first["title"])
            first_content = first["content"] or ""
            for second in rows[index + 1:]:
                title_score = SequenceMatcher(None, first_title, normalize(second["title"])).ratio()
                content_score = text_similarity(first_content, second["content"] or "")
                score = max(title_score, content_score)
                if score >= 0.76:
                    found.append((score, first, second))
        found.sort(key=lambda item: item[0], reverse=True)

        dialog = tk.Toplevel(self.root)
        dialog.title("איתור מתכונים וקבצים דומים")
        dialog.geometry("900x560")
        dialog.transient(self.root)
        frame = ttk.Frame(dialog, padding=15, style="App.TFrame")
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="פריטים דומים לבדיקה", style="Section.TLabel").pack(fill="x", anchor="e")
        ttk.Label(
            frame, text="הזיהוי מבוסס על דמיון בשם ובתוכן. אין מחיקה אוטומטית.",
            style="Muted.TLabel"
        ).pack(fill="x", pady=(3, 10))
        columns = ("score", "second", "first")
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        tree.heading("score", text="דמיון")
        tree.heading("second", text="פריט דומה")
        tree.heading("first", text="פריט מקורי")
        tree.column("score", width=90, anchor="center")
        tree.column("second", width=340, anchor="e")
        tree.column("first", width=340, anchor="e")
        tree.pack(fill="both", expand=True)
        if not found:
            tree.insert("", "end", values=("—", "לא נמצאו התאמות דומות", ""))
        else:
            for score, first, second in found[:250]:
                tree.insert(
                    "", "end",
                    values=(f"{score * 100:.0f}%", second["title"], first["title"])
                )
        ttk.Button(frame, text="סגירה", command=dialog.destroy).pack(anchor="w", pady=(10, 0))

    def open_recycle_bin(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("סל המחזור")
        dialog.geometry("850x520")
        dialog.transient(self.root)
        frame = ttk.Frame(dialog, padding=15, style="App.TFrame")
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="סל המחזור", style="Section.TLabel").pack(fill="x", anchor="e")
        ttk.Label(
            frame, text="קבצים שהוסרו מהספרייה נשארים כאן עד לשחזור או למחיקה סופית.",
            style="Muted.TLabel"
        ).pack(fill="x", pady=(3, 10))
        tree = ttk.Treeview(frame, columns=("deleted", "kind", "title"), show="headings", selectmode="browse")
        tree.heading("deleted", text="תאריך הסרה")
        tree.heading("kind", text="סוג קובץ")
        tree.heading("title", text="שם הקובץ")
        tree.column("deleted", width=150, anchor="e")
        tree.column("kind", width=130, anchor="e")
        tree.column("title", width=500, anchor="e")
        tree.pack(fill="both", expand=True)
        rows = self.library.recycled_rows()
        item_map = {}
        for row in rows:
            iid = str(row["id"])
            item_map[iid] = row
            tree.insert("", "end", iid=iid, values=(row["deleted_at"], file_kind(row["extension"]), row["title"]))

        controls = ttk.Frame(frame, style="App.TFrame")
        controls.pack(fill="x", pady=(10, 0))

        def restore():
            selected = tree.selection()
            if not selected:
                return
            self.library.restore_from_recycle_bin(int(selected[0]))
            self.refresh()
            dialog.destroy()
            self.status_value.set("הקובץ שוחזר מסל המחזור.")

        def delete_forever():
            selected = tree.selection()
            if not selected:
                return
            row = item_map[selected[0]]
            if messagebox.askyesno(
                "מחיקה סופית",
                f"למחוק לצמיתות את ״{row['title']}״? לא ניתן יהיה לשחזר אותו.",
                parent=dialog
            ):
                self.library.permanently_delete(int(selected[0]))
                tree.delete(selected[0])
                item_map.pop(selected[0], None)

        ttk.Button(controls, text="שחזור פריט", style="Primary.TButton", command=restore).pack(side="right")
        ttk.Button(controls, text="מחיקה לצמיתות", style="Danger.TButton", command=delete_forever).pack(side="right", padx=6)
        ttk.Button(controls, text="סגירה", command=dialog.destroy).pack(side="left")

    def restore_backup(self):
        source = filedialog.askopenfilename(
            title="בחירת קובץ גיבוי לשחזור",
            filetypes=[("גיבוי ספרייה (ZIP)", "*.zip")]
        )
        if not source:
            return
        if not messagebox.askyesno(
            "אישור שחזור גיבוי",
            "השחזור יחליף את הספרייה הנוכחית בתוכן הגיבוי. מומלץ ליצור גיבוי עדכני לפני ההמשך.\nלהמשיך?"
        ):
            return
        try:
            safety_backup = self.library.backup_dir / (
                "גיבוי-בטיחות-לפני-שחזור-" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".zip"
            )
            self.library.create_backup(safety_backup)
            count = self.library.restore_backup(Path(source))
            self.refresh()
            messagebox.showinfo(
                "השחזור הושלם",
                f"הספרייה שוחזרה בהצלחה. נטענו {count} רשומות.\n"
                f"גיבוי בטיחות של הספרייה הקודמת נשמר כאן:\n{safety_backup}"
            )
            self.status_value.set("הספרייה שוחזרה מגיבוי. גיבוי בטיחות של המצב הקודם נשמר.")
        except Exception as exc:
            messagebox.showerror("שחזור נכשל", f"הספרייה הנוכחית נשמרה ככל האפשר.\n{exc}")

    def _cleanup_preview_temp(self):
        if self.preview_temp_directory is not None:
            shutil.rmtree(self.preview_temp_directory, ignore_errors=True)
            self.preview_temp_directory = None

    def add_files(self):
        paths = filedialog.askopenfilenames(title="בחירת קבצים להוספה", filetypes=[("כל הקבצים", "*.*")])
        if paths:
            self.import_paths([Path(p) for p in paths])

    def add_folder(self):
        folder = filedialog.askdirectory(title="בחירת תיקייה לייבוא")
        if folder:
            root_folder = Path(folder)
            paths = [p for p in root_folder.rglob("*") if p.is_file()]
            self.import_paths(paths, source_root=root_folder)

    def on_drop(self, event):
        try:
            paths = [Path(p) for p in self.root.tk.splitlist(event.data)]
            expanded, roots = [], []
            for path in paths:
                if path.is_dir():
                    roots.append(path.resolve())
                    expanded.extend(p for p in path.rglob("*") if p.is_file())
                elif path.is_file():
                    expanded.append(path)
            source_root = None
            if roots:
                try:
                    source_root = Path(os.path.commonpath([str(root) for root in roots]))
                except ValueError:
                    source_root = roots[0] if len(roots) == 1 else None
            self.import_paths(expanded, source_root=source_root)
        except Exception as exc:
            messagebox.showerror("שגיאה", f"לא ניתן לקרוא את הפריטים שנגררו.\n{exc}")

    def import_paths(self, paths, source_root: Path | None = None):
        paths = [Path(path) for path in paths]
        if not paths:
            return
        added, skipped, failed = 0, 0, 0
        home = self.library.home.resolve()
        progress_window = tk.Toplevel(self.root)
        progress_window.title("ייבוא קבצים")
        progress_window.geometry("460x150")
        progress_window.transient(self.root)
        progress_window.protocol("WM_DELETE_WINDOW", lambda: None)
        progress_window.grab_set()
        frame = ttk.Frame(progress_window, padding=18, style="App.TFrame")
        frame.pack(fill="both", expand=True)
        progress_text = tk.StringVar(value=f"מכין לייבוא {len(paths)} קבצים…")
        ttk.Label(frame, textvariable=progress_text, style="Panel.TLabel", anchor="e").pack(fill="x", pady=(0, 10))
        progress = ttk.Progressbar(frame, mode="determinate", maximum=len(paths), value=0)
        progress.pack(fill="x")
        progress_window.update_idletasks()
        try:
            for index, path in enumerate(paths, 1):
                progress_text.set(f"קובץ {index} מתוך {len(paths)}: {path.name[:55]}")
                progress.configure(value=index - 1)
                progress_window.update_idletasks()
                try:
                    path = path.resolve()
                    if home == path or home in path.parents:
                        skipped += 1
                    else:
                        relative_folder = ""
                        if source_root is not None:
                            try:
                                relative_folder = path.parent.relative_to(Path(source_root).resolve()).as_posix()
                            except (ValueError, OSError):
                                relative_folder = ""
                        ok, _message = self.library.import_one(path, relative_folder=relative_folder)
                        if ok:
                            added += 1
                        else:
                            skipped += 1
                except Exception:
                    failed += 1
                progress.configure(value=index)
                progress_window.update_idletasks()
        finally:
            try:
                progress_window.grab_release()
                progress_window.destroy()
            except tk.TclError:
                pass
        self.refresh()
        self.status_value.set(f"הייבוא הסתיים: {added} נוספו, {skipped} דולגו, {failed} נכשלו.")
        if failed:
            messagebox.showwarning("סיום ייבוא", f"נוספו {added} קבצים.\nדולגו {skipped} קבצים.\n{failed} קבצים לא נקלטו.")

    def open_selected(self):
        row = self.selected_row()
        if not row:
            messagebox.showinfo("פתיחת קובץ", "יש לבחור קובץ מהרשימה.")
            return
        self.library.record_opened(row["id"])
        path = Path(row["stored_path"])
        self.show_viewer(path, row["title"], row["content"])
        if self.sort_value.get() == "נצפו לאחרונה":
            self.refresh()

    def show_viewer(self, path: Path, title: str, content: str = ""):
        """Show files inside the application; Office files are converted headlessly to PDF."""
        window = tk.Toplevel(self.root)
        window.title(title)
        window.geometry("1320x920")
        window.minsize(820, 620)
        window.configure(background=self.colors["background"])
        window.transient(self.root)
        header = ttk.Frame(window, padding=(16, 12), style="Panel.TFrame")
        header.pack(fill="x", padx=12, pady=(12, 4))
        ttk.Label(header, text=title, style="Header.TLabel", anchor="e").pack(fill="x")
        body = ttk.Frame(window, padding=10, style="App.TFrame")
        body.pack(fill="both", expand=True)

        suffix = path.suffix.lower()
        preview_path = path
        temporary_directory = None
        notice = ""
        if suffix in OFFICE_TYPES:
            preview_path, temporary_directory, notice = convert_office_to_pdf(path)

        preview_document = None
        if suffix in IMAGE_TYPES and Image is not None:
            self._show_image_in_viewer(body, path)
        elif (suffix == ".pdf" or suffix in EBOOK_TYPES or preview_path != path) and fitz is not None and Image is not None:
            preview_document = self._show_pdf_in_viewer(body, preview_path, window)
        else:
            if notice:
                ttk.Label(
                    body,
                    text=notice,
                    foreground="#9A4B18",
                    background="#FFF4E5",
                    padding=10,
                    wraplength=950,
                    justify="right",
                ).pack(fill="x", pady=(0, 8))
            text = content or extract_text(path)
            if not text:
                text = "לא ניתן להציג כרגע תצוגה חזותית לקובץ הזה. הקובץ נשאר שמור בספרייה."
            self._show_text_in_viewer(body, text)

        def close_preview():
            try:
                if preview_document is not None:
                    preview_document.close()
            except Exception:
                pass
            if temporary_directory is not None:
                shutil.rmtree(temporary_directory, ignore_errors=True)
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", close_preview)

    def _show_text_in_viewer(self, parent, text: str):
        text_frame = ttk.Frame(parent)
        text_frame.pack(fill="both", expand=True)
        viewer = tk.Text(
            text_frame, wrap="word", font=("Segoe UI", 12), padx=18, pady=16,
            background="#FFFFFF", foreground=self.colors["text"],
            insertbackground=self.colors["accent"], relief="flat",
            highlightthickness=1, highlightbackground=self.colors["border"]
        )
        scroll = ttk.Scrollbar(text_frame, orient="vertical", command=viewer.yview)
        viewer.configure(yscrollcommand=scroll.set)
        viewer.pack(side="right", fill="both", expand=True)
        scroll.pack(side="left", fill="y")
        viewer.insert("1.0", text)
        viewer.tag_configure("rtl", justify="right")
        viewer.tag_add("rtl", "1.0", "end")
        viewer.configure(state="disabled")

    def _show_pdf_in_viewer(self, parent, path: Path, window):
        """Render PDF/ebook pages lazily so large books do not exhaust memory."""
        if fitz is None or Image is None or ImageTk is None:
            ttk.Label(parent, text="רכיב תצוגת העמודים אינו זמין במחשב זה.").pack(pady=20)
            return None

        import io

        try:
            document = fitz.open(path)
            if document.page_count < 1:
                raise ValueError("המסמך אינו מכיל עמודים.")

            canvas = tk.Canvas(parent, background="#E8EDF3", highlightthickness=0)
            scrollbar = ttk.Scrollbar(parent, orient="vertical")
            inner = ttk.Frame(canvas)
            inner.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            scrollbar.configure(command=canvas.yview)
            canvas.pack(side="right", fill="both", expand=True)
            scrollbar.pack(side="left", fill="y")

            page_frames = []
            for number, page in enumerate(document, 1):
                page_width = max(1.0, float(page.rect.width))
                page_height = max(1.0, float(page.rect.height))
                scale = min(1.35, 1120.0 / page_width)
                frame_height = int(page_height * scale) + 42
                frame = ttk.Frame(inner, width=1160, height=frame_height, style="Panel.TFrame")
                frame.pack(fill="x", padx=8, pady=7)
                frame.pack_propagate(False)
                ttk.Label(frame, text=f"טעינת עמוד {number}…", style="Panel.TLabel").pack(expand=True)
                page_frames.append((frame, scale, frame_height))

            state = {"scheduled": False, "photos": {}}

            def schedule_render(_event=None):
                if not state["scheduled"]:
                    state["scheduled"] = True
                    window.after_idle(render_visible_pages)

            def render_visible_pages():
                state["scheduled"] = False
                try:
                    inner.update_idletasks()
                    bounds = canvas.bbox("all")
                    if not bounds:
                        return
                    visible_top = canvas.canvasy(0)
                    visible_bottom = canvas.canvasy(max(1, canvas.winfo_height()))
                    wanted = set()
                    for index, (frame, _scale, frame_height) in enumerate(page_frames):
                        top = frame.winfo_y()
                        bottom = top + frame_height
                        if bottom >= visible_top - 1100 and top <= visible_bottom + 1100:
                            wanted.add(index)

                    for index in sorted(wanted):
                        if index in state["photos"]:
                            continue
                        frame, scale, _height = page_frames[index]
                        try:
                            pixmap = document[index].get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                            image = Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")
                            photo = ImageTk.PhotoImage(image)
                            for child in frame.winfo_children():
                                child.destroy()
                            image_label = ttk.Label(frame, image=photo, style="Panel.TLabel")
                            image_label.image = photo
                            image_label.pack(expand=True)
                            state["photos"][index] = photo
                        except Exception as page_error:
                            for child in frame.winfo_children():
                                child.destroy()
                            ttk.Label(
                                frame,
                                text=f"לא ניתן לטעון את עמוד {index + 1}: {page_error}",
                                style="Panel.TLabel",
                                wraplength=1080,
                            ).pack(expand=True)
                    # Keep only nearby pages in memory. Distant pages are redrawn when revisited.
                    keep = set()
                    for index in wanted:
                        keep.update(range(max(0, index - 1), min(len(page_frames), index + 2)))
                    for index in list(state["photos"]):
                        if index not in keep:
                            frame = page_frames[index][0]
                            for child in frame.winfo_children():
                                child.destroy()
                            ttk.Label(frame, text=f"עמוד {index + 1} — גלילה לטעינה", style="Panel.TLabel").pack(expand=True)
                            state["photos"].pop(index, None)
                except tk.TclError:
                    return

            def on_scroll(*args):
                scrollbar.set(*args)
                schedule_render()

            scrollbar.configure(command=lambda *args: (canvas.yview(*args), schedule_render()))
            canvas.configure(yscrollcommand=on_scroll)
            canvas.bind("<Configure>", schedule_render)
            inner.bind("<Configure>", lambda _event: (canvas.configure(scrollregion=canvas.bbox("all")), schedule_render()))
            window.after_idle(schedule_render)
            return document
        except Exception as exc:
            ttk.Label(parent, text=f"לא ניתן להציג את המסמך: {exc}", wraplength=900).pack(pady=20)
            return None

    def _show_image_in_viewer(self, parent, path: Path):
        try:
            parent.update_idletasks()
            width = max(1100, parent.winfo_width() - 70)
            height = max(760, parent.winfo_height() - 70)
            with Image.open(path) as source:
                source.seek(0)
                image = source.copy()
            resampling = getattr(Image, "Resampling", Image)
            image.thumbnail((width, height), getattr(resampling, "LANCZOS", 1))
            photo = ImageTk.PhotoImage(image)
            label = ttk.Label(parent, image=photo, style="Panel.TLabel")
            label.image = photo
            label.pack(expand=True, fill="both")
        except Exception as exc:
            ttk.Label(parent, text=f"לא ניתן להציג את התמונה: {exc}").pack()


        try:
            image = Image.open(path)
            image.seek(0)
            image.thumbnail((900, 650))
            photo = ImageTk.PhotoImage(image.copy())
            label = ttk.Label(parent, image=photo)
            label.image = photo
            label.pack(expand=True)
        except Exception as exc:
            ttk.Label(parent, text=f"לא ניתן להציג את התמונה: {exc}").pack()

    def rename_selected(self):
        row = self.selected_row()
        if not row:
            messagebox.showinfo("שינוי שם", "יש לבחור קובץ מהרשימה.")
            return
        value = simpledialog.askstring("שינוי שם", "הקלידו את השם החדש:", initialvalue=row["title"], parent=self.root)
        if value and value.strip():
            try:
                self.library.rename(row["id"], value)
                self.refresh()
                self.status_value.set("שם הקובץ עודכן בהצלחה.")
            except Exception as exc:
                messagebox.showerror("שגיאה", f"לא ניתן לשנות את שם הקובץ.\n{exc}")

    def delete_selected(self):
        row = self.selected_row()
        if not row:
            messagebox.showinfo("הסרה", "יש לבחור קובץ מהרשימה.")
            return
        if messagebox.askyesno("אישור הסרה", f"האם להעביר את ״{row['title']}״ לסל המחזור?\nאפשר יהיה לשחזר אותו בהמשך."):
            self.library.move_to_recycle_bin(row["id"])
            self.refresh()
            self.status_value.set("הקובץ הועבר לסל המחזור.")

    def open_data_folder(self):
        try:
            os.startfile(str(self.library.home))
        except Exception:
            messagebox.showinfo("תיקיית הנתונים", str(self.library.home))

    def backup(self):
        default_name = "גיבוי-הספרייה-" + datetime.now().strftime("%Y%m%d-%H%M") + ".zip"
        target = filedialog.asksaveasfilename(
            title="שמירת גיבוי", defaultextension=".zip",
            initialfile=default_name,
            filetypes=[("קובץ גיבוי דחוס", "*.zip")]
        )
        if not target:
            return
        try:
            self.library.create_backup(Path(target))
            self.status_value.set("הגיבוי נוצר בהצלחה.")
            messagebox.showinfo("גיבוי הושלם", "הגיבוי נשמר בהצלחה.")
        except Exception as exc:
            messagebox.showerror("שגיאה ביצירת גיבוי", str(exc))

    def settings(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("הגדרות")
        dialog.geometry("620x570")
        dialog.minsize(540, 480)
        dialog.transient(self.root)
        dialog.grab_set()
        frame = ttk.Frame(dialog, padding=18, style="App.TFrame")
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="הגדרות הספרייה", style="Section.TLabel").pack(fill="x", pady=(0, 12))

        theme_var = tk.StringVar(value=self.library.get_setting("color_theme", "כחול"))
        font_var = tk.StringVar(value=self.library.get_setting("font_size", "10"))
        zoom_var = tk.StringVar(value=self.library.get_setting("preview_zoom", "1.0"))
        auto_var = tk.BooleanVar(value=self.library.get_setting("auto_backup_enabled", "0") == "1")
        days_var = tk.StringVar(value=self.library.get_setting("auto_backup_days", "7"))

        def option_row(title, variable, values, width=18):
            row = ttk.Frame(frame, style="App.TFrame")
            row.pack(fill="x", pady=5)
            ttk.Label(row, text=title, style="Panel.TLabel").pack(side="right", padx=(0, 10))
            ttk.Combobox(
                row, textvariable=variable, values=values,
                state="readonly", width=width, justify="right"
            ).pack(side="right")
            return row

        option_row("ערכת צבעים:", theme_var, ["כחול", "ירוק", "סגול"])
        option_row("גודל גופן:", font_var, ["9", "10", "11", "12", "13", "14"])
        option_row("גודל תצוגה מקדימה:", zoom_var, ["0.75", "1.0", "1.25", "1.5"])
        ttk.Checkbutton(
            frame, text="הפעלת גיבוי אוטומטי",
            variable=auto_var
        ).pack(anchor="e", pady=(12, 4))
        option_row("מרווח בין גיבויים אוטומטיים:", days_var, ["1", "3", "7", "14", "30"], width=10)
        ttk.Label(
            frame, text="הגיבויים האוטומטיים נשמרים בתיקיית ״גיבויים״ ונשמרים חמשת האחרונים.",
            style="Muted.TLabel", wraplength=560, justify="right"
        ).pack(fill="x", pady=(4, 10))
        ttk.Label(frame, text="מיקום שמירת הנתונים במחשב:", style="Panel.TLabel").pack(fill="x", anchor="e")
        ttk.Label(
            frame, text=str(self.library.home), style="Muted.TLabel",
            wraplength=560, justify="right"
        ).pack(fill="x", pady=5)

        def save_settings():
            try:
                days = int(days_var.get())
                font = int(font_var.get())
                zoom = float(zoom_var.get())
                if days < 1 or font < 8 or font > 15 or zoom < 0.5 or zoom > 1.5:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("הגדרות", "אחד הערכים אינו תקין.", parent=dialog)
                return
            self.library.set_setting("color_theme", theme_var.get())
            self.library.set_setting("font_size", str(font))
            self.library.set_setting("preview_zoom", str(zoom))
            was_enabled = self.library.get_setting("auto_backup_enabled", "0") == "1"
            self.library.set_setting("auto_backup_enabled", "1" if auto_var.get() else "0")
            self.library.set_setting("auto_backup_days", str(days))
            if auto_var.get() and not was_enabled and not self.library.get_setting("last_auto_backup", ""):
                self.library.set_setting("last_auto_backup", datetime.now().isoformat(timespec="seconds"))
            self._apply_styles()
            self._set_preview_placeholder("בחרו תמונה או קובץ מהרשימה") if not self.preview_path else self._render_selected_preview()
            self.status_value.set("ההגדרות נשמרו.")
            dialog.destroy()

        buttons = ttk.Frame(frame, style="App.TFrame")
        buttons.pack(fill="x", pady=(14, 0))
        ttk.Button(buttons, text="פתיחת תיקיית הנתונים", command=self.open_data_folder).pack(side="right")
        ttk.Button(buttons, text="שמירה והחלה", style="Primary.TButton", command=save_settings).pack(side="right", padx=8)
        ttk.Button(buttons, text="סגירה", command=dialog.destroy).pack(side="left")

    def close(self):
        try:
            if self._auto_backup_job is not None:
                try:
                    self.root.after_cancel(self._auto_backup_job)
                except tk.TclError:
                    pass
            self._cleanup_preview_temp()
            self.library.db.commit()
            self.library.db.close()
        finally:
            self.root.destroy()

    def run(self):
        self.root.mainloop()

def main():
    try:
        App().run()
    except Exception as exc:
        try:
            messagebox.showerror("שגיאה בהפעלת התוכנה", f"התוכנה לא הצליחה להיפתח.\n{exc}")
        except Exception:
            pass
        raise

if __name__ == "__main__":
    main()