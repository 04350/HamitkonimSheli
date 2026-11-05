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
from datetime import datetime
from difflib import SequenceMatcher
from html.parser import HTMLParser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

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
APP_VERSION = "1.3.0"
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

class RecipeLibrary:
    def __init__(self):
        self.home = data_home()
        self.files_dir = self.home / "הקבצים שלי"
        self.db_path = self.home / "הספרייה.sqlite3"
        self.db = sqlite3.connect(self.db_path)
        self.db.row_factory = sqlite3.Row
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
        # Backward-compatible migration: existing libraries keep all their files and gain
        # a favorites flag with a default of off.
        item_columns = {
            row["name"] for row in self.db.execute("PRAGMA table_info(items)").fetchall()
        }
        if "is_favorite" not in item_columns:
            self.db.execute(
                "ALTER TABLE items ADD COLUMN is_favorite INTEGER NOT NULL DEFAULT 0"
            )
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
        self.db.commit()

    def rows(self, favorite_only: bool = False, collection_id: int | None = None):
        if collection_id is not None:
            return self.db.execute("""
                SELECT items.* FROM items
                INNER JOIN item_collections ON item_collections.item_id = items.id
                WHERE item_collections.collection_id = ?
                ORDER BY items.id DESC
            """, (int(collection_id),)).fetchall()
        if favorite_only:
            return self.db.execute(
                "SELECT * FROM items WHERE is_favorite = 1 ORDER BY id DESC"
            ).fetchall()
        return self.db.execute("SELECT * FROM items ORDER BY id DESC").fetchall()

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

    def import_one(self, source: Path) -> tuple[bool, str]:
        if not source.is_file():
            return False, "זה אינו קובץ."
        if source.stat().st_size > MAX_FILE_SIZE:
            return False, "הקובץ גדול מדי. ניתן לייבא קבצים עד 250 מגה־בייט."
        digest = file_digest(source)
        if self.db.execute("SELECT id FROM items WHERE digest=?", (digest,)).fetchone():
            return False, "הקובץ כבר נמצא בספרייה."
        content = extract_text(source)
        title = choose_title(source, content)
        target = unique_path(self.files_dir, title, source.suffix.lower())
        shutil.copy2(source, target)
        try:
            self.db.execute(
                "INSERT INTO items(title,original_name,stored_path,extension,content,digest,imported_at) VALUES(?,?,?,?,?,?,?)",
                (title, source.name, str(target), source.suffix.lower(), content, digest, datetime.now().strftime("%d/%m/%Y %H:%M"))
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
        new = unique_path(self.files_dir, title, old.suffix)
        old.rename(new)
        self.db.execute("UPDATE items SET title=?,stored_path=? WHERE id=?", (safe_title(title), str(new), item_id))
        self.db.commit()

class App:
    def __init__(self):
        self.library = RecipeLibrary()
        self.root = ROOT_CLASS()
        self.root.title(APP_NAME)
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
        self._filter_options = {}
        self.favorite_button = None
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

    def _apply_styles(self):
        """Apply a modern, high-contrast desktop theme."""
        self.colors = {
            "background": "#EEF3F9",
            "surface": "#FFFFFF",
            "navy": "#172A45",
            "hero": "#10243D",
            "text": "#25364D",
            "muted": "#718198",
            "border": "#DCE5EF",
            "accent": "#087E8B",
            "accent_hover": "#066571",
            "soft": "#E6F3F5",
            "row_alt": "#F5F8FC",
            "selected": "#DCEFF2",
            "danger": "#B83E4B",
        }
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        c = self.colors
        self.root.configure(background=c["background"])
        self.style.configure(".", font=("Segoe UI", 10), foreground=c["text"])
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
            "TButton", font=("Segoe UI", 9, "bold"), padding=(14, 10),
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
            ("▤  פתיחת תיקיית הנתונים", self.open_data_folder),
            ("↥  יצירת גיבוי", self.backup),
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
        ttk.Button(
            filter_frame, text="📁  ניהול אוספים",
            command=self.open_collections_dialog
        ).pack(side="left", padx=(4, 0))

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
            preview, text="פתיחת הקובץ  ↗", command=self.open_selected,
            style="Primary.TButton"
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

        table_body = ttk.Frame(table_card, style="Card.TFrame")
        table_body.pack(fill="both", expand=True)
        columns = ("favorite", "date", "kind", "title")
        self.tree = ttk.Treeview(table_body, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("favorite", text="★", anchor="center")
        self.tree.heading("date", text="תאריך הוספה", anchor="e")
        self.tree.heading("kind", text="סוג קובץ", anchor="e")
        self.tree.heading("title", text="שם המתכון או הקובץ", anchor="e")
        self.tree.column("favorite", width=46, minwidth=42, anchor="center", stretch=False)
        self.tree.column("date", width=145, minwidth=125, anchor="e", stretch=False)
        self.tree.column("kind", width=115, minwidth=95, anchor="e", stretch=False)
        self.tree.column("title", width=510, minwidth=230, anchor="e", stretch=True)
        scroll = ttk.Scrollbar(table_body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.tag_configure("even", background=self.colors["surface"])
        self.tree.tag_configure("odd", background=self.colors["row_alt"])
        self.tree.pack(side="right", fill="both", expand=True)
        scroll.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.show_selected)
        self.tree.bind("<Double-1>", lambda _event: self.open_selected())

        status_frame = ttk.Frame(outer, padding=(10, 7), style="Card.TFrame")
        status_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(status_frame, textvariable=self.status_value, style="Status.TLabel", anchor="e").pack(fill="x")

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
        if current not in options:
            current = "כל הקבצים"
        self.filter_value.set(current)

    def _on_filter_change(self, _event=None):
        self.refresh()

    def refresh(self):
        selected = self.tree.selection()
        selected_id = selected[0] if selected else None
        self._update_filter_options()
        mode, filter_id = self._filter_options.get(self.filter_value.get(), ("all", None))
        rows = self.library.rows(
            favorite_only=(mode == "favorites"),
            collection_id=(filter_id if mode == "collection" else None)
        )
        query = self.search_value.get()
        for item in self.tree.get_children():
            self.tree.delete(item)
        ranked = []
        for row in rows:
            score = relevance(row["title"], row["content"], query)
            if score > 0:
                ranked.append((score, row))
        ranked.sort(key=lambda pair: (pair[0], pair[1]["id"]), reverse=True)
        self.current_rows = {}
        for index, (_score, row) in enumerate(ranked):
            iid = str(row["id"])
            self.current_rows[iid] = row
            ext = file_kind(row["extension"])
            star = "★" if row["is_favorite"] else ""
            self.tree.insert(
                "", "end", iid=iid,
                values=(star, row["imported_at"], ext, row["title"]),
                tags=("even" if index % 2 == 0 else "odd",)
            )
        self.status_value.set(f"נמצאו {len(ranked)} קבצים בתצוגה הנוכחית.")
        if selected_id and selected_id in self.current_rows:
            self.tree.selection_set(selected_id)
            self.tree.focus(selected_id)
            self.show_selected()
        else:
            self.preview_path = None
            self.preview_photo = None
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
        selected = self.tree.selection()
        return self.current_rows.get(selected[0]) if selected else None

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
        path = Path(row["stored_path"])
        collection_names = self.library.item_collection_names(row["id"])
        collection_text = ", ".join(collection_names) if collection_names else "ללא אוסף"
        info = (
            f"שם: {row['title']}\n"
            f"קובץ מקורי: {row['original_name']}\n"
            f"נוסף בתאריך: {row['imported_at']}\n"
            f"אוספים: {collection_text}"
        )
        self.detail_label.configure(text=info)
        self._update_selection_actions(row)
        self.preview_photo = None
        self.preview_path = path if path.suffix.lower() in IMAGE_TYPES and Image is not None else None
        if self.preview_path is not None:
            self._set_preview_placeholder("טוען תצוגה מקדימה…")
            self.root.after_idle(self._render_selected_image)
            return
        self._set_preview_placeholder("אין תצוגה מקדימה לקובץ הזה")

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
        self._preview_resize_job = self.root.after(140, self._render_selected_image)

    def _render_selected_image(self):
        self._preview_resize_job = None
        path = self.preview_path
        if path is None or not path.is_file() or Image is None or ImageTk is None:
            return
        try:
            self.preview_image.update_idletasks()
            width = self.preview_image.winfo_width()
            height = self.preview_image.winfo_height()
            if width < 80 or height < 80:
                return
            target_size = (width - 28, height - 28)
            with Image.open(path) as source:
                source.seek(0)
                image = source.convert("RGBA")
            resampling = getattr(Image, "Resampling", Image)
            image.thumbnail(target_size, getattr(resampling, "LANCZOS", 1))
            self.preview_photo = ImageTk.PhotoImage(image)
            self.preview_image.delete("all")
            self.preview_image.create_image(
                width // 2, height // 2,
                image=self.preview_photo, anchor="center",
                tags=("preview-image",)
            )
        except Exception as exc:
            self.preview_photo = None
            self._set_preview_placeholder(f"לא ניתן להציג את התמונה\n{exc}")

    def add_files(self):
        paths = filedialog.askopenfilenames(title="בחירת קבצים להוספה", filetypes=[("כל הקבצים", "*.*")])
        if paths:
            self.import_paths([Path(p) for p in paths])

    def add_folder(self):
        folder = filedialog.askdirectory(title="בחירת תיקייה לייבוא")
        if folder:
            paths = [p for p in Path(folder).rglob("*") if p.is_file()]
            self.import_paths(paths)

    def on_drop(self, event):
        try:
            paths = [Path(p) for p in self.root.tk.splitlist(event.data)]
            expanded = []
            for path in paths:
                if path.is_dir():
                    expanded.extend(p for p in path.rglob("*") if p.is_file())
                elif path.is_file():
                    expanded.append(path)
            self.import_paths(expanded)
        except Exception as exc:
            messagebox.showerror("שגיאה", f"לא ניתן לקרוא את הפריטים שנגררו.\n{exc}")

    def import_paths(self, paths):
        added, skipped, failed = 0, 0, 0
        home = self.library.home.resolve()
        for path in paths:
            try:
                path = path.resolve()
                if home == path or home in path.parents:
                    skipped += 1
                    continue
                ok, _message = self.library.import_one(path)
                if ok:
                    added += 1
                else:
                    skipped += 1
            except Exception:
                failed += 1
        self.refresh()
        self.status_value.set(f"הייבוא הסתיים: {added} קבצים נוספו, {skipped} דולגו, {failed} נכשלו.")
        if failed:
            messagebox.showwarning("סיום ייבוא", f"נוספו {added} קבצים.\nדולגו {skipped} קבצים.\n{failed} קבצים לא נקלטו.")

    def open_selected(self):
        row = self.selected_row()
        if not row:
            messagebox.showinfo("פתיחת קובץ", "יש לבחור קובץ מהרשימה.")
            return
        path = Path(row["stored_path"])
        self.show_viewer(path, row["title"], row["content"])

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
        if messagebox.askyesno("אישור הסרה", f"האם להסיר את ״{row['title']}״ מהספרייה?\nהקובץ יימחק מתיקיית הנתונים המקומית."):
            self.library.remove(row["id"])
            self.refresh()
            self.status_value.set("הקובץ הוסר.")

    def open_data_folder(self):
        try:
            os.startfile(str(self.library.home))
        except Exception:
            messagebox.showinfo("תיקיית הנתונים", str(self.library.home))

    def backup(self):
        default_name = "גיבוי-הספרייה-" + datetime.now().strftime("%Y%m%d-%H%M") + ".zip"
        target = filedialog.asksaveasfilename(title="שמירת גיבוי", defaultextension=".zip", initialfile=default_name, filetypes=[("קובץ גיבוי דחוס", "*.zip")])
        if not target:
            return
        try:
            self.library.db.commit()
            with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.write(self.library.db_path, "הספרייה.sqlite3")
                for path in self.library.files_dir.iterdir():
                    if path.is_file():
                        archive.write(path, "הקבצים שלי/" + path.name)
            self.status_value.set("הגיבוי נוצר בהצלחה.")
            messagebox.showinfo("גיבוי הושלם", "הגיבוי נשמר בהצלחה.")
        except Exception as exc:
            messagebox.showerror("שגיאה ביצירת גיבוי", str(exc))

    def settings(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("הגדרות")
        dialog.geometry("480x250")
        dialog.transient(self.root)
        dialog.grab_set()
        frame = ttk.Frame(dialog, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="הגדרות הספרייה", font=("Arial", 16, "bold"), anchor="e").pack(fill="x", pady=(0, 12))
        ttk.Label(frame, text="מיקום שמירת הנתונים במחשב:", anchor="e").pack(fill="x")
        ttk.Label(frame, text=str(self.library.home), wraplength=430, justify="right").pack(fill="x", pady=5)
        ttk.Label(frame, text="הנתונים נשמרים במחשב בלבד. התוכנה אינה מתחברת לאינטרנט ואינה שולחת מידע.", wraplength=430, justify="right").pack(fill="x", pady=8)
        ttk.Button(frame, text="פתיחת תיקיית הנתונים", command=self.open_data_folder).pack(anchor="e", pady=5)
        ttk.Button(frame, text="סגירה", command=dialog.destroy).pack(anchor="e")

    def close(self):
        try:
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
