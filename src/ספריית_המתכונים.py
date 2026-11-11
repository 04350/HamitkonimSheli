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
from collections import OrderedDict
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from html.parser import HTMLParser
import tkinter as tk
import threading
import queue
import math
from tkinter import ttk, filedialog, messagebox, simpledialog

from recipe_center import RecipeCenter, recipe_search_text, compile_shopping_list, build_recipe_book_docx, export_recipe_book_pdf, parse_recipe_text, is_recipe_record, recipe_missing_parts
from recipe_scanner import recognize_image, best_title_from_pdf_page

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
    from PIL import Image, ImageTk, ExifTags, ImageOps
except Exception:
    Image = ImageTk = ExifTags = ImageOps = None
try:
    import pytesseract
except Exception:
    pytesseract = None

APP_NAME = "ספריית המתכונים שלי"
APP_VERSION = "2.0.0"
def _rounded_points(x1, y1, x2, y2, radius=14, steps=8):
    """Create a polygon outline with smooth rounded corners."""
    width = max(1, x2 - x1)
    height = max(1, y2 - y1)
    radius = max(0, min(radius, width / 2, height / 2))
    if radius <= 0:
        return [x1, y1, x2, y1, x2, y2, x1, y2]
    points = []
    corners = [
        (x1 + radius, y1 + radius, 180, 270),
        (x2 - radius, y1 + radius, 270, 360),
        (x2 - radius, y2 - radius, 0, 90),
        (x1 + radius, y2 - radius, 90, 180),
    ]
    for cx, cy, start, end in corners:
        for step in range(steps + 1):
            angle = math.radians(start + (end - start) * step / steps)
            points.extend((cx + radius * math.cos(angle), cy + radius * math.sin(angle)))
    return points


class ModernButton(tk.Canvas):
    """A lightweight rounded button drawn on a canvas rather than a native Windows rectangle."""
    def __init__(
        self, parent, text, command=None, *, width=128, height=42,
        bg="#E7EEE7", fg="#294238", hover_bg="#DDE8DD",
        active_bg="#345646", active_fg="#FFFFFF", radius=13,
        font=("Segoe UI", 10, "bold"), align="center",
        canvas_bg=None, disabled=False, padx=15,
    ):
        self.label_text = str(text)
        self.command = command
        self.fill = bg
        self.foreground = fg
        self.hover_fill = hover_bg or bg
        self.pressed_fill = active_bg or bg
        self.active_foreground = active_fg
        self.radius = radius
        self.font_spec = font
        self.align = align
        self.disabled = bool(disabled)
        self.hovered = False
        self.pressed = False
        self.selected = False
        self.padx = padx
        super().__init__(
            parent, width=width, height=height,
            bg=canvas_bg or parent.cget("background"),
            highlightthickness=0, bd=0, relief="flat",
            cursor="arrow" if disabled else "hand2",
        )
        self.bind("<Configure>", self._draw)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self._draw()

    def _on_enter(self, _event=None):
        if not self.disabled:
            self.hovered = True
            self._draw()

    def _on_leave(self, _event=None):
        self.hovered = False
        self.pressed = False
        self._draw()

    def _on_press(self, _event=None):
        if not self.disabled:
            self.pressed = True
            self._draw()

    def _on_release(self, event=None):
        was_pressed = self.pressed
        self.pressed = False
        self._draw()
        if was_pressed and not self.disabled and self.command:
            width, height = max(1, self.winfo_width()), max(1, self.winfo_height())
            if event is None or (0 <= event.x <= width and 0 <= event.y <= height):
                self.command()

    def set_active(self, active=True):
        self.selected = bool(active)
        self._draw()

    def set_text(self, text):
        self.label_text = str(text)
        self._draw()

    def set_disabled(self, disabled=True):
        self.disabled = bool(disabled)
        self.configure(cursor="arrow" if self.disabled else "hand2")
        self._draw()

    def _draw(self, _event=None):
        self.delete("all")
        width, height = max(2, self.winfo_width()), max(2, self.winfo_height())
        fill, foreground = self.fill, self.foreground
        if self.disabled:
            fill, foreground = "#ECEFEA", "#A1AAA2"
        elif self.selected or self.pressed:
            fill, foreground = self.pressed_fill, self.active_foreground
        elif self.hovered:
            fill = self.hover_fill
        self.create_polygon(
            *_rounded_points(1, 1, width - 1, height - 1, self.radius),
            fill=fill, outline=fill, width=1, smooth=True, splinesteps=18
        )
        if self.align == "right":
            x, anchor = width - self.padx, "e"
        elif self.align == "left":
            x, anchor = self.padx, "w"
        else:
            x, anchor = width / 2, "center"
        self.create_text(
            x, height / 2, text=self.label_text, fill=foreground,
            font=self.font_spec, anchor=anchor, justify="right",
            tags=("button-label",)
        )


class RoundedPanel(tk.Canvas):
    """Canvas-backed rounded surface with a soft shadow and a normal Tk frame inside."""
    def __init__(
        self, parent, *, surface="#FFFFFF", background="#F5F1E9",
        outline="#E7E8E1", shadow="#E9E6DD", radius=20, padding=2,
        width=300, height=220, show_shadow=True,
    ):
        self.surface_color = surface
        self.outline_color = outline
        self.shadow_color = shadow
        self.radius = radius
        self.inner_padding = padding
        self.content_inset = max(padding + 1, min(10, max(5, radius // 2)))
        self.show_shadow = show_shadow
        super().__init__(
            parent, width=width, height=height, bg=background,
            highlightthickness=0, bd=0, relief="flat"
        )
        self.content = tk.Frame(self, bg=surface, bd=0, highlightthickness=0)
        self.window_id = self.create_window(
            self.content_inset, self.content_inset, anchor="nw", window=self.content
        )
        self.bind("<Configure>", self._resize)
        self._draw()

    def _resize(self, event=None):
        width = max(4, event.width if event else self.winfo_width())
        height = max(4, event.height if event else self.winfo_height())
        pad = self.content_inset
        self.coords(self.window_id, pad, pad)
        self.itemconfigure(
            self.window_id,
            width=max(1, width - pad * 2),
            height=max(1, height - pad * 2),
        )
        self._draw(event)

    def _draw(self, event=None):
        self.delete("panel-shape")
        width = max(4, event.width if event else self.winfo_width())
        height = max(4, event.height if event else self.winfo_height())
        if self.show_shadow:
            self.create_polygon(
                *_rounded_points(3, 4, width - 2, height - 1, self.radius),
                fill=self.shadow_color, outline=self.shadow_color,
                tags=("panel-shape",), smooth=True, splinesteps=18
            )
        self.create_polygon(
            *_rounded_points(1, 1, width - 4, height - 4, self.radius),
            fill=self.surface_color, outline=self.outline_color,
            width=1, tags=("panel-shape",), smooth=True, splinesteps=18
        )
        self.tag_lower("panel-shape")

    def set_outline(self, color):
        self.outline_color = color
        self._draw()

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
GALLERY_THUMBNAIL_WIDTH = 320
GALLERY_THUMBNAIL_HEIGHT = 145
GALLERY_CACHE_MEMORY_BUDGET_BYTES = 96 * 1024 * 1024
GALLERY_THUMBNAIL_ESTIMATED_BYTES = GALLERY_THUMBNAIL_WIDTH * GALLERY_THUMBNAIL_HEIGHT * 4
GALLERY_CACHE_MAX_IMAGES = max(1, GALLERY_CACHE_MEMORY_BUDGET_BYTES // GALLERY_THUMBNAIL_ESTIMATED_BYTES)

def _portable_libreoffice_archive_version(archive_path: Path) -> str:
    """Read the tiny build marker stored inside the bundled Office ZIP."""
    try:
        with zipfile.ZipFile(archive_path, "r") as archive:
            try:
                value = archive.read(".recipe-library-version").decode("utf-8", errors="replace").strip()
                return value or "legacy"
            except KeyError:
                return "legacy"
    except (OSError, zipfile.BadZipFile):
        return "legacy"


def _extract_portable_libreoffice(archive_path: Path, progress_callback=None) -> Path | None:
    """Extract or update the bundled Office runtime; cached versions are left untouched."""
    base = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local"))
    runtime_root = base / APP_NAME / "רכיבים"
    target = runtime_root / "LibreOffice"
    bundle_version = _portable_libreoffice_archive_version(archive_path)
    marker_name = ".recipe-library-version"
    marker_path = target / marker_name
    try:
        existing_version = marker_path.read_text(encoding="utf-8").strip()
    except (OSError, UnicodeError):
        existing_version = ""

    for executable_name in ("soffice.com", "soffice.exe"):
        existing = target / "program" / executable_name
        if existing.is_file() and existing_version == bundle_version:
            return existing

    runtime_root.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix="LibreOffice-extract-", dir=runtime_root))
    try:
        with zipfile.ZipFile(archive_path, "r") as archive:
            members = archive.infolist()
            if len(members) > 30000:
                raise ValueError("ארכיון LibreOffice מכיל מספר חריג של קבצים.")
            total_size = sum(info.file_size for info in members)
            if total_size > 3 * 1024 * 1024 * 1024:
                raise ValueError("ארכיון LibreOffice גדול מהמגבלה המותרת.")
            root_resolved = staging.resolve()
            total_members = len(members)
            for index, info in enumerate(members, 1):
                name = info.filename.replace("\\", "/")
                parts = [part for part in name.split("/") if part not in ("", ".")]
                if not parts or name.startswith("/") or any(part == ".." or ":" in part for part in parts):
                    raise ValueError("ארכיון LibreOffice מכיל נתיב שאינו בטוח.")
                destination = staging.joinpath(*parts)
                if not destination.resolve().is_relative_to(root_resolved):
                    raise ValueError("ארכיון LibreOffice מנסה לכתוב מחוץ לתיקיית החילוץ.")
                if info.is_dir():
                    destination.mkdir(parents=True, exist_ok=True)
                else:
                    if ((info.external_attr >> 16) & 0o170000) == 0o120000:
                        raise ValueError("ארכיון LibreOffice מכיל קישור סמלי שאינו נתמך.")
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    with archive.open(info, "r") as source, destination.open("wb") as output:
                        shutil.copyfileobj(source, output, length=1024 * 1024)

                if progress_callback and (index % 100 == 0 or index == total_members):
                    if progress_callback("extracting", index, total_members) is False:
                        raise InterruptedError("חילוץ LibreOffice בוטל.")

        executable = next(
            (staging / "program" / name for name in ("soffice.com", "soffice.exe")
             if (staging / "program" / name).is_file()),
            None,
        )
        if executable is None:
            raise ValueError("לא נמצא soffice בתוך ארכיון LibreOffice.")
        (staging / marker_name).write_text(bundle_version, encoding="utf-8")
        if target.exists():
            shutil.rmtree(target, ignore_errors=True)
        os.replace(staging, target)
        return next(
            (target / "program" / name for name in ("soffice.com", "soffice.exe")
             if (target / "program" / name).is_file()),
            None,
        )
    except Exception:
        shutil.rmtree(staging, ignore_errors=True)
        raise

def find_libreoffice(progress_callback=None) -> Path | None:
    """Find a folder, a version-matched portable ZIP archive, or a system installation."""
    names = ("soffice.com", "soffice.exe")
    roots: list[Path] = []
    if getattr(sys, "frozen", False):
        roots.append(Path(sys.executable).resolve().parent)
    bundle_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    roots.extend((bundle_dir, Path(__file__).resolve().parent, Path(__file__).resolve().parent.parent))

    for root in roots:
        for name in names:
            candidate = root / "LibreOffice" / "program" / name
            if candidate.is_file():
                return candidate

    # The distributed ZIP has only one LibreOffice.zip, not thousands of loose files.
    for root in roots:
        archive_path = root / "LibreOffice.zip"
        if archive_path.is_file():
            if progress_callback:
                progress_callback("checking", 0, 0)
            return _extract_portable_libreoffice(
                archive_path, progress_callback=progress_callback
            )

    for root in (Path(os.environ.get("PROGRAMFILES", "C:/Program Files")),
                 Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)"))):
        for name in names:
            candidate = root / "LibreOffice" / "program" / name
            if candidate.is_file():
                return candidate
    for name in names + ("soffice",):
        located = shutil.which(name)
        if located:
            return Path(located)
    return None


def convert_office_to_pdf(
    source: Path,
    timeout: int = 180,
    progress_callback=None,
) -> tuple[Path | None, Path | None, str]:
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
    executable = find_libreoffice(progress_callback=progress_callback)
    if executable is None:
        return None, None, "מנוע התצוגה של Office לא נמצא. יש לוודא שתיקיית LibreOffice נמצאת לצד קובץ התוכנה."

    temporary_directory = Path(tempfile.mkdtemp(prefix="ספריית-מתכונים-office-"))
    output_directory = temporary_directory / "output"
    profile_directory = temporary_directory / "profile"
    output_directory.mkdir()
    profile_directory.mkdir()
    output_pdf = output_directory / (source.stem + ".pdf")
    if progress_callback and progress_callback("converting", 0, 0) is False:
        return None, None, "הפעולה בוטלה."
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

OCR_TITLE_CACHE: dict[str, str] = {}
OCR_FILE_CACHE: dict[str, dict] = {}
_LAST_OCR_TITLE = ""


def ocr_image(image) -> str:
    """Improved local OCR with a title candidate derived from printed text size."""
    global _LAST_OCR_TITLE
    _LAST_OCR_TITLE = ""
    if pytesseract is None:
        return ""
    result = recognize_image(image, pytesseract)
    _LAST_OCR_TITLE = str(result.get("title") or "")
    return str(result.get("text") or "")


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
    global _LAST_OCR_TITLE
    suffix = path.suffix.lower()
    cache_key = str(path.resolve())
    cached = OCR_FILE_CACHE.pop(cache_key, None)
    if cached:
        if cached.get("title"):
            OCR_TITLE_CACHE[cache_key] = str(cached["title"])
        return str(cached.get("text") or "")[:MAX_TEXT_LENGTH]
    OCR_TITLE_CACHE.pop(cache_key, None)
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
                first_page_title = ""
                for page_index, page in enumerate(pdf):
                    page_text = page.get_text("text").strip()
                    if page_index == 0:
                        first_page_title = best_title_from_pdf_page(page)
                    if not page_text and Image is not None:
                        pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2), alpha=False)
                        import io
                        with Image.open(io.BytesIO(pix.tobytes("png"))) as page_image:
                            _LAST_OCR_TITLE = ""
                            page_text = ocr_image(page_image)
                        if page_index == 0 and _LAST_OCR_TITLE:
                            first_page_title = _LAST_OCR_TITLE
                    pages.append(page_text)
                if first_page_title:
                    OCR_TITLE_CACHE[cache_key] = first_page_title
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
                    _LAST_OCR_TITLE = ""
                    recognized = ocr_image(image)
            if _LAST_OCR_TITLE:
                OCR_TITLE_CACHE[cache_key] = _LAST_OCR_TITLE
            text = "\n".join(part for part in (metadata, recognized) if part)
        else:
            text = ""
    except Exception:
        text = ""
    return (text or "")[:MAX_TEXT_LENGTH]


def choose_title(path: Path, content: str) -> str:
    cache_key = str(path.resolve())
    visual_title = OCR_TITLE_CACHE.pop(cache_key, "")
    if not visual_title and path.suffix.lower() in IMAGE_TYPES and Image is not None and pytesseract is not None:
        try:
            with Image.open(path) as image:
                visual_title = str(recognize_image(image.convert("RGB"), pytesseract).get("title") or "")
        except Exception:
            visual_title = ""
    if visual_title:
        return safe_title(visual_title)

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


def _recipe_data_from_row(row) -> dict:
    """Decode recipe metadata already selected with the item row, without another SQL query."""
    try:
        value = json.loads(row["recipe_data"] or "{}")
        return value if isinstance(value, dict) else {}
    except (KeyError, TypeError, ValueError, IndexError):
        return {}


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
        self._recipe_options_cache = None
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
        self.db.execute("CREATE INDEX IF NOT EXISTS idx_items_active_id ON items(is_deleted, id DESC)")
        self.db.execute("CREATE INDEX IF NOT EXISTS idx_items_favorite_id ON items(is_deleted, is_favorite, id DESC)")
        self.db.execute("CREATE INDEX IF NOT EXISTS idx_item_collections_collection_item ON item_collections(collection_id, item_id)")
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
        self._invalidate_recipe_options()

    def restore_from_recycle_bin(self, item_id: int):
        self.db.execute(
            "UPDATE items SET is_deleted=0, deleted_at='' WHERE id=? AND is_deleted=1",
            (int(item_id),)
        )
        self.db.commit()
        self._invalidate_recipe_options()

    def permanently_delete(self, item_id: int):
        row = self.db.execute(
            "SELECT stored_path FROM items WHERE id=? AND is_deleted=1", (int(item_id),)
        ).fetchone()
        if row:
            Path(row["stored_path"]).unlink(missing_ok=True)
            self.db.execute("DELETE FROM items WHERE id=?", (int(item_id),))
            self.db.commit()
            self._invalidate_recipe_options()

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
        self._invalidate_recipe_options()

    def _invalidate_recipe_options(self):
        self._recipe_options_cache = None

    def _recipe_option_values(self):
        if self._recipe_options_cache is None:
            folders = set()
            categories = set()
            for row in self.db.execute("SELECT recipe_data FROM items WHERE is_deleted=0"):
                try:
                    data = json.loads(row["recipe_data"] or "{}")
                    if not isinstance(data, dict):
                        continue
                    folder = str(data.get("source_folder") or "").strip("/")
                    category = str(data.get("category") or "").strip()
                except (ValueError, AttributeError, TypeError):
                    continue
                if category:
                    categories.add(category)
                if folder:
                    parts = folder.split("/")
                    for index in range(1, len(parts) + 1):
                        folders.add("/".join(parts[:index]))
            self._recipe_options_cache = (
                tuple(sorted(categories, key=str.casefold)),
                tuple(sorted(folders, key=lambda value: (value.count("/"), value.casefold()))),
            )
        return self._recipe_options_cache

    def all_source_folders(self) -> list[str]:
        return list(self._recipe_option_values()[1])

    def tag_names(self, item_id: int) -> list[str]:
        rows = self.db.execute("""
            SELECT tags.name FROM tags
            INNER JOIN item_tags ON item_tags.tag_id=tags.id
            WHERE item_tags.item_id=? ORDER BY tags.name COLLATE NOCASE
        """, (int(item_id),)).fetchall()
        return [row["name"] for row in rows]

    def tags_by_item(self, item_ids) -> dict[int, list[str]]:
        """Load tags for many items in a few queries instead of one query per recipe."""
        ids = list(dict.fromkeys(int(item_id) for item_id in item_ids))
        result = {item_id: [] for item_id in ids}
        for start in range(0, len(ids), 800):
            chunk = ids[start:start + 800]
            if not chunk:
                continue
            placeholders = ",".join("?" for _ in chunk)
            rows = self.db.execute(
                f"SELECT item_tags.item_id, tags.name FROM item_tags "
                f"INNER JOIN tags ON tags.id=item_tags.tag_id "
                f"WHERE item_tags.item_id IN ({placeholders}) "
                "ORDER BY tags.name COLLATE NOCASE",
                chunk,
            ).fetchall()
            for row in rows:
                result[int(row["item_id"])].append(row["name"])
        return result

    def all_tags(self) -> list[str]:
        return [row["name"] for row in self.db.execute("SELECT name FROM tags ORDER BY name COLLATE NOCASE")]

    def all_categories(self) -> list[str]:
        return list(self._recipe_option_values()[0])

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
                self._invalidate_recipe_options()
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

    def prepare_import(
        self,
        source: Path,
        relative_folder: str = "",
        known_digests: set[str] | None = None,
    ) -> dict:
        """Read/hash/parse a source file without touching SQLite; safe for a worker thread."""
        source = Path(source).resolve()
        if not source.is_file():
            raise ValueError("זה אינו קובץ.")
        if source.stat().st_size > MAX_FILE_SIZE:
            raise ValueError("הקובץ גדול מדי. ניתן לייבא קבצים עד 250 מגה־בייט.")
        digest = file_digest(source)
        if known_digests is not None:
            if digest in known_digests:
                return {"duplicate": True, "source": source, "digest": digest}
            known_digests.add(digest)

        content = extract_text(source)
        title = choose_title(source, content)
        safe_parts = []
        for part in Path(str(relative_folder or "")).parts:
            if part in ("", ".", "..", "/", "\\"):
                continue
            cleaned = safe_title(part)
            if cleaned and cleaned not in (".", ".."):
                safe_parts.append(cleaned)
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
        return {
            "duplicate": False,
            "source": source,
            "digest": digest,
            "content": content,
            "title": title,
            "safe_parts": safe_parts,
            "recipe_data": recipe_data,
        }

    def import_prepared(self, prepared: dict) -> tuple[bool, str]:
        """Commit a prepared item. This method must run on the UI/SQLite thread."""
        if prepared.get("duplicate"):
            return False, "הקובץ כבר נמצא בספרייה."
        source = Path(prepared["source"])
        digest = str(prepared["digest"])
        if self.db.execute("SELECT id FROM items WHERE digest=?", (digest,)).fetchone():
            return False, "הקובץ כבר נמצא בספרייה."
        safe_parts = list(prepared.get("safe_parts") or [])
        target_folder = self.files_dir.joinpath(*safe_parts)
        target_folder.mkdir(parents=True, exist_ok=True)
        title = safe_title(prepared.get("title") or source.stem)
        target = unique_path(target_folder, title, source.suffix.lower())
        shutil.copy2(source, target)
        try:
            self.db.execute(
                "INSERT INTO items(title,original_name,stored_path,extension,content,digest,imported_at,recipe_data) VALUES(?,?,?,?,?,?,?,?)",
                (title, source.name, str(target), source.suffix.lower(),
                 str(prepared.get("content") or ""), digest,
                 datetime.now().strftime("%d/%m/%Y %H:%M"),
                 json.dumps(prepared.get("recipe_data") or {}, ensure_ascii=False))
            )
            self.db.commit()
            self._invalidate_recipe_options()
        except Exception:
            target.unlink(missing_ok=True)
            raise
        return True, title

    def import_one(self, source: Path, relative_folder: str = "") -> tuple[bool, str]:
        prepared = self.prepare_import(source, relative_folder=relative_folder)
        return self.import_prepared(prepared)


    def remove(self, item_id: int):
        row = self.db.execute("SELECT stored_path FROM items WHERE id=?", (item_id,)).fetchone()
        if row:
            Path(row["stored_path"]).unlink(missing_ok=True)
            self.db.execute("DELETE FROM items WHERE id=?", (item_id,))
            self.db.commit()
            self._invalidate_recipe_options()

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
        self.root.geometry("1680x960")
        self.root.minsize(1120, 720)
        self.style = ttk.Style(self.root)
        self._apply_styles()
        self.search_value = tk.StringVar()
        self.filter_value = tk.StringVar(value="כל הקבצים")
        self.type_filter_value = tk.StringVar(value="כל הסוגים")
        self.category_filter_value = tk.StringVar(value="כל הקטגוריות")
        self.tag_filter_value = tk.StringVar(value="כל התגיות")
        self.sort_value = tk.StringVar(value="תאריך: חדש לישן")
        self.view_mode = tk.StringVar(value="גלריה")
        self.current_page = "recipes"
        self.nav_buttons = {}
        self._embedded_pages = []
        self.reader_preview_mode = tk.StringVar(
            value=self.library.get_setting("reader_preview_mode", "image")
        )
        self._gallery_limit = 60
        self.reader_photos = []
        self._scan_thread = None
        self._import_cancel_event = None
        self.selected_item_id = None
        self.gallery_photos = []
        # Dynamic LRU: sized from the library's image count, with a memory budget.
        self._gallery_image_cache = OrderedDict()
        self._gallery_image_cache_limit = 0
        self.date_from_value = tk.StringVar()
        self.date_to_value = tk.StringVar()
        self.folder_filter_value = tk.StringVar(value="כל התיקיות")
        self.recipe_filter_value = tk.StringVar(value="מתכונים בלבד")
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
        """Premium warm cookbook theme with consistent hierarchy and calm surfaces."""
        theme = self.library.get_setting("color_theme", "שמנת")
        palettes = {
            "שמנת": {
                "background": "#F4F1E9", "surface": "#FFFFFF", "navy": "#263D32",
                "hero": "#20392F", "text": "#304238", "muted": "#89938B",
                "border": "#E9E6DE", "accent": "#B97853", "accent_hover": "#A96543",
                "row_alt": "#F8F7F2", "selected": "#E9F0E8", "danger": "#B44F52",
                "sidebar": "#1E352B", "sidebar_active": "#355747",
                "cover": "#E9E7DC", "shadow": "#E8E4DA", "soft": "#EDF2EA",
                "pill": "#EEF3EC", "pill_text": "#48654F", "sidebar_text": "#DCE6DC",
            },
            "מרווה": {
                "background": "#EEF2EC", "surface": "#FFFFFF", "navy": "#294238",
                "hero": "#29483A", "text": "#30463B", "muted": "#7D8A80",
                "border": "#DDE6DC", "accent": "#7B987C", "accent_hover": "#668468",
                "row_alt": "#F6F9F4", "selected": "#E5EEE4", "danger": "#B44F52",
                "sidebar": "#203A30", "sidebar_active": "#3A5B47",
                "cover": "#DDE9DD", "shadow": "#E0E9DF", "soft": "#E7EFE5",
                "pill": "#EAF1E8", "pill_text": "#49694F", "sidebar_text": "#DCE7DC",
            },
            "לילה": {
                "background": "#151F1A", "surface": "#202D25", "navy": "#E7EFE7",
                "hero": "#17241D", "text": "#E2EAE1", "muted": "#A0ADA2",
                "border": "#3A4B40", "accent": "#B99170", "accent_hover": "#CAA27F",
                "row_alt": "#29372E", "selected": "#33483A", "danger": "#E18D8D",
                "sidebar": "#101A15", "sidebar_active": "#354E3E",
                "cover": "#314237", "shadow": "#0F1712", "soft": "#2A3A30",
                "pill": "#304337", "pill_text": "#D6E4D5", "sidebar_text": "#D5E1D6",
            },
            "כחול": {
                "background": "#F0F3F7", "surface": "#FFFFFF", "navy": "#23364D",
                "hero": "#1C3048", "text": "#304157", "muted": "#8390A0",
                "border": "#E1E7ED", "accent": "#318B91", "accent_hover": "#26767C",
                "row_alt": "#F7F9FB", "selected": "#E4EFF0", "danger": "#B83E4B",
                "sidebar": "#1D3348", "sidebar_active": "#35566B",
                "cover": "#E2EBEF", "shadow": "#E2E8EE", "soft": "#EAF1F2",
                "pill": "#E8F1F2", "pill_text": "#3A6F72", "sidebar_text": "#DCE6ED",
            },
            "ירוק": {
                "background": "#EEF4F1", "surface": "#FFFFFF", "navy": "#1E4238",
                "hero": "#153A30", "text": "#2C463D", "muted": "#7E9087",
                "border": "#DDE9E2", "accent": "#16846D", "accent_hover": "#116B58",
                "row_alt": "#F5F9F6", "selected": "#E0EFE7", "danger": "#B83E4B",
                "sidebar": "#173C32", "sidebar_active": "#2B5949",
                "cover": "#DDEDE4", "shadow": "#DFEAE3", "soft": "#E4F0E8",
                "pill": "#E1F0E8", "pill_text": "#27745F", "sidebar_text": "#D9E8DF",
            },
            "סגול": {
                "background": "#F3F0F7", "surface": "#FFFFFF", "navy": "#3C3153",
                "hero": "#302542", "text": "#443A54", "muted": "#8A8196",
                "border": "#E6E0EC", "accent": "#8063A6", "accent_hover": "#684D8D",
                "sidebar": "#302542", "sidebar_active": "#50406A",
                "cover": "#EAE3F1", "shadow": "#E7E0EC", "soft": "#EEE8F4",
                "pill": "#EEE8F4", "pill_text": "#6E568C", "row_alt": "#F9F7FB",
                "selected": "#EEE8F4", "danger": "#B83E4B", "sidebar_text": "#E8E1EF",
            },
        }
        self.colors = dict(palettes.get(theme, palettes["שמנת"]))
        try:
            font_size = max(9, min(15, int(self.library.get_setting("font_size", "10"))))
        except (ValueError, TypeError):
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
        self.style.configure("Status.TLabel", background=c["background"], foreground=c["muted"], padding=(10, 7), font=("Segoe UI", 9))
        self.style.configure(
            "TButton", font=("Segoe UI", max(9, font_size - 1), "bold"),
            padding=(15, 10), background=c["soft"], foreground=c["navy"],
            bordercolor=c["border"], focuscolor=c["accent"], relief="flat"
        )
        self.style.map(
            "TButton",
            background=[("disabled", c["row_alt"]), ("pressed", c["selected"]), ("active", c["selected"])],
            foreground=[("disabled", c["muted"]), ("active", c["navy"])]
        )
        self.style.configure(
            "Primary.TButton", background=c["accent"], foreground="#FFFFFF",
            bordercolor=c["accent"], padding=(17, 11)
        )
        self.style.map(
            "Primary.TButton",
            background=[("disabled", c["muted"]), ("pressed", c["accent_hover"]), ("active", c["accent_hover"])],
            foreground=[("disabled", "#F5FAFA"), ("!disabled", "#FFFFFF")]
        )
        self.style.configure(
            "Danger.TButton", background="#FCEDEF", foreground=c["danger"],
            bordercolor="#F3D4D8", padding=(12, 9)
        )
        self.style.map("Danger.TButton", background=[("pressed", "#F4D9DD"), ("active", "#F8E4E7")])
        self.style.configure(
            "TEntry", fieldbackground=c["row_alt"], foreground=c["navy"],
            padding=(13, 11), bordercolor=c["border"], insertcolor=c["navy"],
            font=("Segoe UI", 11), relief="flat"
        )
        self.style.configure(
            "TCombobox", fieldbackground=c["row_alt"], background=c["surface"],
            foreground=c["navy"], arrowcolor=c["muted"],
            padding=(10, 9), bordercolor=c["border"], relief="flat"
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", c["row_alt"])],
            selectbackground=[("readonly", c["row_alt"])],
            selectforeground=[("readonly", c["navy"])],
            foreground=[("readonly", c["navy"])]
        )
        self.style.configure(
            "Treeview", rowheight=44, font=("Segoe UI", 10),
            background=c["surface"], fieldbackground=c["surface"],
            foreground=c["text"], bordercolor=c["border"], borderwidth=0
        )
        self.style.configure(
            "Treeview.Heading", font=("Segoe UI", 9, "bold"),
            background=c["row_alt"], foreground=c["muted"],
            padding=(13, 12), relief="flat"
        )
        self.style.map("Treeview", background=[("selected", c["selected"])], foreground=[("selected", c["navy"])])
        self.style.map("Treeview.Heading", background=[("active", c["soft"])])
        self.style.configure(
            "Vertical.TScrollbar", background=c["border"],
            troughcolor=c["background"], bordercolor=c["background"], arrowsize=14
        )


    def _push_embedded_page(self, title=""):
        """Show a full-size screen inside the existing application window."""
        if not hasattr(self, "_embedded_pages"):
            self._embedded_pages = []
        if self._embedded_pages:
            try:
                self._embedded_pages[-1].place_forget()
            except tk.TclError:
                pass
        host = getattr(self, "workspace", self.shell)
        page = tk.Frame(host, bg=self.colors["background"], bd=0, highlightthickness=0)
        page.place(relx=0, rely=0, relwidth=1, relheight=1)
        page.lift()
        page.page_title = title
        self._embedded_pages.append(page)
        return page

    def _pop_embedded_page(self, page=None):
        """Close the current embedded screen and reveal the previous screen."""
        stack = getattr(self, "_embedded_pages", [])
        if not stack:
            return
        current = stack[-1]
        if page is not None and current is not page:
            try:
                page.destroy()
            except tk.TclError:
                pass
            try:
                stack.remove(page)
            except ValueError:
                pass
            return
        stack.pop()
        try:
            current.destroy()
        except tk.TclError:
            pass
        if stack:
            previous = stack[-1]
            try:
                previous.place(relx=0, rely=0, relwidth=1, relheight=1)
                previous.lift()
            except tk.TclError:
                pass
        else:
            try:
                self.workspace.lift()
            except (tk.TclError, AttributeError):
                try:
                    self.shell.lift()
                except tk.TclError:
                    pass

    def _embedded_header(self, parent, title, subtitle="", back_command=None, back_text="→  חזרה לספרייה"):
        """Create the shared in-app page header with a clear back action."""
        header = tk.Frame(parent, bg=self.colors["background"], padx=22, pady=16)
        header.pack(fill="x")
        title_area = tk.Frame(header, bg=self.colors["background"])
        title_area.pack(side="right", fill="x", expand=True)
        tk.Label(
            title_area, text=title, bg=self.colors["background"],
            fg=self.colors["navy"], font=("Segoe UI", 22, "bold"), anchor="e"
        ).pack(fill="x")
        if subtitle:
            tk.Label(
                title_area, text=subtitle, bg=self.colors["background"],
                fg=self.colors["muted"], font=("Segoe UI", 9), anchor="e"
            ).pack(fill="x", pady=(3, 0))
        if back_command is None:
            back_command = lambda target=parent: self._pop_embedded_page(target)
        ModernButton(
            header, text=back_text,
            command=back_command,
            width=150, height=40, bg=self.colors["surface"],
            fg=self.colors["navy"], hover_bg=self.colors["soft"],
            active_bg=self.colors["selected"], radius=13,
            font=("Segoe UI", 9, "bold"), canvas_bg=self.colors["background"]
        ).pack(side="left", padx=(0, 5))
        return header

    def _save_reader_sash_position(self, _event=None):
        try:
            width = max(1, self.content_splitter.winfo_width())
            sash_x = self.content_splitter.sash_coord(0)[0]
            reader_width = max(300, width - sash_x)
            self.library.set_setting("reader_pane_width", str(reader_width))
        except (tk.TclError, ValueError, IndexError):
            pass

    def _reset_reader_sash_position(self, _event=None):
        """Double-click the divider to restore a balanced reader width."""
        try:
            width = max(1, self.content_splitter.winfo_width())
            reader_width = 410
            target = max(440, width - reader_width)
            self.content_splitter.sash_place(0, target, 0)
            self.library.set_setting("reader_pane_width", str(reader_width))
        except (tk.TclError, ValueError, IndexError):
            pass
        return "break"

    def _set_reader_preview_mode(self, mode):
        if mode not in ("image", "text"):
            return
        self.reader_preview_mode.set(mode)
        self.library.set_setting("reader_preview_mode", mode)
        self._update_reader_mode_buttons()
        self._render_reader(self.selected_row())

    def _update_reader_mode_buttons(self):
        if not hasattr(self, "reader_image_button"):
            return
        active = self.reader_preview_mode.get()
        self.reader_image_button.set_active(active == "image")
        self.reader_text_button.set_active(active == "text")

    def _build(self):
        """Modern, airy recipe-first home with rounded surfaces and quiet navigation."""
        c = self.colors
        self.shell = tk.Frame(self.root, bg=c["background"], bd=0)
        self.shell.pack(fill="both", expand=True)

        sidebar = tk.Frame(self.shell, bg=c["sidebar"], width=226, bd=0)
        sidebar.pack(side="right", fill="y")
        sidebar.pack_propagate(False)

        brand = tk.Frame(sidebar, bg=c["sidebar"], padx=22, pady=24)
        brand.pack(fill="x")
        mark = tk.Canvas(brand, width=48, height=48, bg=c["sidebar"], highlightthickness=0, bd=0)
        mark.pack(anchor="e")
        mark.create_polygon(*_rounded_points(2, 2, 46, 46, 14), fill=c["accent"], outline=c["accent"])
        mark.create_text(24, 24, text="מ", fill="#FFFFFF", font=("Segoe UI", 19, "bold"))
        tk.Label(
            brand, text="ספריית המתכונים", bg=c["sidebar"], fg="#FFFFFF",
            font=("Segoe UI", 15, "bold"), anchor="e"
        ).pack(fill="x", pady=(12, 2))
        tk.Label(
            brand, text="כל הטעמים שלך, במקום אחד", bg=c["sidebar"],
            fg="#B8C8BC", font=("Segoe UI", 9), anchor="e"
        ).pack(fill="x")
        tk.Frame(sidebar, bg="#3C5748", height=1).pack(fill="x", padx=20, pady=(3, 18))

        self.nav_buttons = {}
        tk.Label(sidebar, text="הספרייה", bg=c["sidebar"], fg="#91A99A",
                 font=("Segoe UI", 9, "bold"), anchor="e", padx=22).pack(fill="x", pady=(0, 7))
        for key, glyph, label in [
            ("recipes", "▦", "המתכונים שלי"),
            ("favorites", "★", "מועדפים"),
            ("files", "▤", "כל הקבצים"),
        ]:
            self._add_nav_button(sidebar, key, glyph, label)

        tk.Label(sidebar, text="כלים", bg=c["sidebar"], fg="#91A99A",
                 font=("Segoe UI", 9, "bold"), anchor="e", padx=22).pack(fill="x", pady=(25, 7))
        for key, glyph, label in [
            ("scanner", "⌕", "סריקה חכמה"),
            ("center", "◉", "מרכז המתכונים"),
            ("import", "＋", "הוספת קבצים"),
            ("folder", "▣", "ייבוא תיקייה"),
        ]:
            self._add_nav_button(sidebar, key, glyph, label)
        tk.Frame(sidebar, bg="#3C5748", height=1).pack(fill="x", padx=20, pady=(18, 10))
        for key, glyph, label in [
            ("backup", "↥", "גיבוי ושחזור"),
            ("recycle", "♻", "סל המחזור"),
            ("settings", "⚙", "הגדרות ועיצוב"),
        ]:
            self._add_nav_button(sidebar, key, glyph, label)

        footer = tk.Frame(sidebar, bg=c["sidebar"], padx=20, pady=18)
        footer.pack(side="bottom", fill="x")
        tk.Label(
            footer, text="●  עבודה מקומית ומאובטחת", bg=c["sidebar"],
            fg="#B9C9BD", font=("Segoe UI", 8), anchor="e"
        ).pack(fill="x")
        tk.Label(
            footer, text=f"גרסה {APP_VERSION}", bg=c["sidebar"],
            fg="#81998A", font=("Segoe UI", 8), anchor="e"
        ).pack(fill="x", pady=(4, 0))

        workspace = tk.Frame(self.shell, bg=c["background"], padx=26, pady=23)
        workspace.pack(side="left", fill="both", expand=True)
        self.workspace = workspace
        header = tk.Frame(workspace, bg=c["background"])
        header.pack(fill="x", pady=(0, 17))
        heading = tk.Frame(header, bg=c["background"])
        heading.pack(side="right", fill="x", expand=True)
        self.page_title_label = tk.Label(
            heading, text="המתכונים שלי", bg=c["background"], fg=c["navy"],
            font=("Segoe UI", 27, "bold"), anchor="e"
        )
        self.page_title_label.pack(fill="x")
        self.page_subtitle_label = tk.Label(
            heading, text="אוסף הטעמים שלך — מסודר, זמין ומוכן לקריאה.",
            bg=c["background"], fg=c["muted"], font=("Segoe UI", 10), anchor="e"
        )
        self.page_subtitle_label.pack(fill="x", pady=(5, 0))
        self.count_badge = ModernButton(
            header, text="הספרייה שלך", command=None, width=150, height=38,
            bg=c["pill"], fg=c["pill_text"], hover_bg=c["pill"],
            active_bg=c["pill"], radius=18, font=("Segoe UI", 9, "bold"),
            canvas_bg=c["background"], disabled=True
        )
        self.count_badge.pack(side="left", padx=(0, 5))
        self.count_label = tk.Label(
            header, text="", bg=c["background"], fg=c["muted"],
            font=("Segoe UI", 9, "bold"), anchor="e"
        )
        self.count_label.pack(side="left", padx=(10, 0))
        ModernButton(
            header, text="＋  הוספת מתכון", command=self.add_files,
            width=148, height=42, bg=c["accent"], fg="#FFFFFF",
            hover_bg=c["accent_hover"], active_bg=c["accent_hover"],
            radius=14, font=("Segoe UI", 9, "bold"), canvas_bg=c["background"]
        ).pack(side="left", padx=(0, 12))

        search_shell = RoundedPanel(
            workspace, surface=c["surface"], background=c["background"],
            outline=c["border"], shadow=c["shadow"], radius=18,
            padding=3, height=68, show_shadow=True
        )
        search_shell.pack(fill="x", pady=(0, 13))
        search_inner = tk.Frame(search_shell.content, bg=c["surface"], padx=15, pady=10)
        search_inner.pack(fill="both", expand=True)
        tk.Label(
            search_inner, text="⌕", bg=c["surface"], fg=c["accent"],
            font=("Segoe UI Symbol", 22), width=2
        ).pack(side="right")
        self.search_entry = ttk.Entry(search_inner, textvariable=self.search_value, justify="right")
        self.search_entry.pack(side="right", fill="x", expand=True, ipady=3, padx=(5, 10))
        self.search_entry.bind("<KeyRelease>", lambda _event: self.refresh())
        ModernButton(
            search_inner, text="ניקוי", command=lambda: (self.search_value.set(""), self.refresh()),
            width=76, height=38, bg=c["row_alt"], fg=c["muted"],
            hover_bg=c["soft"], radius=12, font=("Segoe UI", 9, "bold"),
            canvas_bg=c["surface"]
        ).pack(side="left", padx=(2, 0))

        filter_bar = tk.Frame(workspace, bg=c["background"])
        filter_bar.pack(fill="x", pady=(0, 15))
        tk.Label(filter_bar, text="סינון", bg=c["background"], fg=c["muted"],
                 font=("Segoe UI", 9, "bold")).pack(side="right", padx=(0, 7))
        self.category_combo = ttk.Combobox(
            filter_bar, textvariable=self.category_filter_value,
            state="readonly", width=19, justify="right"
        )
        self.category_combo.pack(side="right", padx=(0, 9), ipady=2)
        self.category_combo.bind("<<ComboboxSelected>>", self._on_filter_change)
        self.sort_combo = ttk.Combobox(
            filter_bar, textvariable=self.sort_value, state="readonly", width=18,
            justify="right", values=[
                "תאריך: חדש לישן", "תאריך: ישן לחדש", "שם: א-ת", "שם: ת-א",
                "דירוג: גבוה לנמוך", "נצפו לאחרונה", "סוג קובץ", "גודל: גדול לקטן"
            ]
        )
        self.sort_combo.pack(side="right")
        self.sort_combo.bind("<<ComboboxSelected>>", self._on_filter_change)

        view_controls = tk.Frame(filter_bar, bg=c["background"])
        view_controls.pack(side="left")
        self.view_all_button = ModernButton(
            view_controls, text="☷  רשימה", command=lambda: self._set_view("רשימה"),
            width=104, height=36, bg=c["surface"], fg=c["muted"],
            hover_bg=c["soft"], active_bg=c["sidebar_active"],
            radius=12, font=("Segoe UI", 9, "bold"), canvas_bg=c["background"]
        )
        self.view_all_button.pack(side="left", padx=(0, 5))
        self.view_grid_button = ModernButton(
            view_controls, text="▦  כרטיסים", command=lambda: self._set_view("גלריה"),
            width=112, height=36, bg=c["surface"], fg=c["muted"],
            hover_bg=c["soft"], active_bg=c["sidebar_active"],
            radius=12, font=("Segoe UI", 9, "bold"), canvas_bg=c["background"]
        )
        self.view_grid_button.pack(side="left")

        content_area = tk.Frame(workspace, bg=c["background"])
        content_area.pack(fill="both", expand=True)
        self.content_splitter = tk.PanedWindow(
            content_area, orient="horizontal", background=c["background"],
            sashwidth=13, sashpad=2, sashrelief="raised", borderwidth=0,
            relief="flat", opaqueresize=True, showhandle=True,
            handlesize=20, handlepad=3, sashcursor="sb_h_double_arrow"
        )
        self.content_splitter.pack(fill="both", expand=True)
        self.content_splitter.bind("<ButtonRelease-1>", self._save_reader_sash_position)
        self.content_splitter.bind("<Double-Button-1>", self._reset_reader_sash_position)
        reader_shell = RoundedPanel(
            self.content_splitter, surface=c["surface"], background=c["background"],
            outline=c["border"], shadow=c["shadow"], radius=21,
            padding=2, width=370, show_shadow=True
        )
        self.reader_panel = reader_shell.content
        reader_heading = tk.Frame(self.reader_panel, bg=c["surface"], padx=17, pady=15)
        reader_heading.pack(fill="x")
        tk.Label(
            reader_heading, text="המתכון הנבחר", bg=c["surface"],
            fg=c["navy"], font=("Segoe UI", 14, "bold"), anchor="e"
        ).pack(fill="x")
        tk.Label(
            reader_heading, text="תמונה או טקסט שחולץ · גררו את המפריד לשינוי רוחב",
            bg=c["surface"], fg=c["muted"], font=("Segoe UI", 8), anchor="e"
        ).pack(fill="x", pady=(3, 0))
        tk.Frame(self.reader_panel, bg=c["border"], height=1).pack(fill="x", padx=16, pady=(0, 4))
        preview_modes = tk.Frame(self.reader_panel, bg=c["surface"])
        preview_modes.pack(fill="x", padx=12, pady=(7, 6))
        self.reader_image_button = ModernButton(
            preview_modes, text="תמונה מקורית",
            command=lambda: self._set_reader_preview_mode("image"),
            width=132, height=34, bg=c["row_alt"], fg=c["muted"],
            hover_bg=c["soft"], active_bg=c["sidebar_active"],
            radius=11, font=("Segoe UI", 8, "bold"), canvas_bg=c["surface"]
        )
        self.reader_image_button.pack(side="right", fill="x", expand=True, padx=3)
        self.reader_text_button = ModernButton(
            preview_modes, text="טקסט שחולץ",
            command=lambda: self._set_reader_preview_mode("text"),
            width=132, height=34, bg=c["row_alt"], fg=c["muted"],
            hover_bg=c["soft"], active_bg=c["sidebar_active"],
            radius=11, font=("Segoe UI", 8, "bold"), canvas_bg=c["surface"]
        )
        self.reader_text_button.pack(side="right", fill="x", expand=True, padx=3)

        reader_canvas_frame = tk.Frame(self.reader_panel, bg=c["surface"])
        reader_canvas_frame.pack(fill="both", expand=True, padx=(7, 4), pady=(0, 8))
        self.reader_canvas = tk.Canvas(reader_canvas_frame, bg=c["surface"], highlightthickness=0, bd=0)
        reader_scroll = ttk.Scrollbar(reader_canvas_frame, orient="vertical", command=self.reader_canvas.yview)
        self.reader_canvas.configure(yscrollcommand=reader_scroll.set)
        reader_scroll.pack(side="left", fill="y")
        self.reader_canvas.pack(side="right", fill="both", expand=True)
        self.reader_inner = tk.Frame(self.reader_canvas, bg=c["surface"])
        self.reader_window = self.reader_canvas.create_window((0, 0), window=self.reader_inner, anchor="nw")
        self.reader_inner.bind("<Configure>", lambda _event: self.reader_canvas.configure(scrollregion=self.reader_canvas.bbox("all")))
        self.reader_canvas.bind("<Configure>", lambda event: self.reader_canvas.itemconfigure(self.reader_window, width=event.width))

        catalog_shell = RoundedPanel(
            self.content_splitter, surface=c["surface"], background=c["background"],
            outline=c["border"], shadow=c["shadow"], radius=21,
            padding=2, show_shadow=True
        )
        self.content_splitter.add(catalog_shell, minsize=440, stretch="always")
        self.content_splitter.add(reader_shell, minsize=320, stretch="always")
        def restore_reader_width():
            try:
                saved_width = int(self.library.get_setting("reader_pane_width", "385"))
            except (TypeError, ValueError):
                saved_width = 385
            target = max(440, self.content_splitter.winfo_width() - max(320, min(900, saved_width)))
            self.content_splitter.sash_place(0, target, 0)
        self.content_splitter.after_idle(restore_reader_width)
        catalog = catalog_shell.content
        catalog_header = tk.Frame(catalog, bg=c["surface"], padx=20, pady=16)
        catalog_header.pack(fill="x")
        self.catalog_title = tk.Label(
            catalog_header, text="כל המתכונים", bg=c["surface"],
            fg=c["navy"], font=("Segoe UI", 15, "bold"), anchor="e"
        )
        self.catalog_title.pack(side="right")
        self.catalog_hint = tk.Label(
            catalog_header, text="בחרו מתכון כדי לפתוח אותו לקריאה",
            bg=c["surface"], fg=c["muted"], font=("Segoe UI", 9), anchor="w"
        )
        self.catalog_hint.pack(side="left")
        tk.Frame(catalog, bg=c["border"], height=1).pack(fill="x", padx=17)
        catalog_body = tk.Frame(catalog, bg=c["surface"])
        catalog_body.pack(fill="both", expand=True, padx=11, pady=(7, 12))

        self.list_container = tk.Frame(catalog_body, bg=c["surface"])
        columns = ("favorite", "date", "kind", "size", "title")
        self.tree = ttk.Treeview(self.list_container, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("favorite", text="★", anchor="center")
        self.tree.heading("date", text="תאריך הוספה", anchor="e")
        self.tree.heading("kind", text="סוג קובץ", anchor="e")
        self.tree.heading("size", text="גודל", anchor="e")
        self.tree.heading("title", text="שם המתכון או הקובץ", anchor="e")
        self.tree.column("favorite", width=42, minwidth=38, anchor="center", stretch=False)
        self.tree.column("date", width=125, minwidth=110, anchor="e", stretch=False)
        self.tree.column("kind", width=100, minwidth=85, anchor="e", stretch=False)
        self.tree.column("size", width=82, minwidth=70, anchor="e", stretch=False)
        self.tree.column("title", width=340, minwidth=190, anchor="e", stretch=True)
        scroll = ttk.Scrollbar(self.list_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.tag_configure("even", background=c["surface"])
        self.tree.tag_configure("odd", background=c["row_alt"])
        self.tree.pack(side="right", fill="both", expand=True)
        scroll.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.show_selected)
        self.tree.bind("<Double-1>", lambda _event: self.open_selected())

        self.gallery_container = tk.Frame(catalog_body, bg=c["surface"])
        self.gallery_canvas = tk.Canvas(self.gallery_container, background=c["surface"], highlightthickness=0, bd=0)
        gallery_scroll = ttk.Scrollbar(self.gallery_container, orient="vertical", command=self.gallery_canvas.yview)
        self.gallery_canvas.configure(yscrollcommand=gallery_scroll.set)
        gallery_scroll.pack(side="left", fill="y")
        self.gallery_canvas.pack(side="right", fill="both", expand=True)
        self.gallery_inner = tk.Frame(self.gallery_canvas, bg=c["surface"])
        self.gallery_window = self.gallery_canvas.create_window((0, 0), window=self.gallery_inner, anchor="nw")
        self.gallery_inner.bind("<Configure>", lambda _event: self.gallery_canvas.configure(scrollregion=self.gallery_canvas.bbox("all")))
        self.gallery_canvas.bind("<Configure>", lambda event: self.gallery_canvas.itemconfigure(self.gallery_window, width=event.width))
        self.gallery_canvas.bind("<MouseWheel>", lambda event: self.gallery_canvas.yview_scroll(int(-event.delta / 120), "units"))
        self.gallery_container.pack(fill="both", expand=True)

        # Compatibility controls used by legacy filters and preview commands.
        self._compat_frame = tk.Frame(self.root)
        self.filter_combo = ttk.Combobox(self._compat_frame, textvariable=self.filter_value)
        self.tag_combo = ttk.Combobox(self._compat_frame, textvariable=self.tag_filter_value)
        self.folder_combo = ttk.Combobox(self._compat_frame, textvariable=self.folder_filter_value)
        self.preview_image = tk.Canvas(self._compat_frame, width=1, height=1)
        self.detail_label = ttk.Label(self._compat_frame)
        self.status_frame = tk.Frame(workspace, bg=c["background"])
        self.status_frame.pack(fill="x", pady=(8, 0))
        self.status_label = tk.Label(
            self.status_frame, textvariable=self.status_value,
            bg=c["background"], fg=c["muted"], font=("Segoe UI", 8), anchor="e"
        )
        self.status_label.pack(fill="x")
        self._update_navigation_active()
        self._update_view_buttons()
        self._update_reader_mode_buttons()
        self._render_reader(None)

    def _add_nav_button(self, parent, key, glyph, label):
        button = ModernButton(
            parent, text=f"{glyph}    {label}",
            command=lambda value=key: self._navigate(value),
            width=190, height=43, bg=self.colors["sidebar"],
            fg=self.colors["sidebar_text"], hover_bg=self.colors["sidebar_active"],
            active_bg=self.colors["sidebar_active"], active_fg="#FFFFFF",
            radius=12, font=("Segoe UI", 10, "bold"), align="right",
            canvas_bg=self.colors["sidebar"], padx=18
        )
        button.pack(fill="x", padx=12, pady=3)
        self.nav_buttons[key] = button

    def _update_navigation_active(self):
        active = getattr(self, "current_page", "recipes")
        for key, button in getattr(self, "nav_buttons", {}).items():
            try:
                button.set_active(key == active)
            except (tk.TclError, AttributeError):
                pass

    def _update_view_buttons(self):
        if not hasattr(self, "view_all_button"):
            return
        self.view_all_button.set_active(self.view_mode.get() == "רשימה")
        self.view_grid_button.set_active(self.view_mode.get() == "גלריה")

    def _navigate(self, key):
        if key == "scanner":
            self.open_ocr_scanner()
            return
        if key == "center":
            self.open_recipe_center()
            return
        if key == "import":
            self.add_files()
            return
        if key == "folder":
            self.add_folder()
            return
        if key == "backup":
            self.backup()
            return
        if key == "recycle":
            self.open_recycle_bin()
            return
        if key == "settings":
            self.settings()
            return

        if key in ("recipes", "favorites", "files") and self._embedded_pages:
            while self._embedded_pages:
                self._pop_embedded_page()

        self.current_page = key
        if key == "recipes":
            self.filter_value.set("כל הקבצים")
            self.recipe_filter_value.set("מתכונים בלבד")
            self.page_title_label.configure(text="המתכונים שלי")
            self.page_subtitle_label.configure(text="אוסף הטעמים שלך — מסודר, זמין ומוכן לקריאה.")
            self.catalog_title.configure(text="כל המתכונים")
        elif key == "favorites":
            self.filter_value.set("★ מועדפים")
            self.recipe_filter_value.set("מתכונים בלבד")
            self.page_title_label.configure(text="המועדפים שלי")
            self.page_subtitle_label.configure(text="המתכונים ששמרת קרוב ללב, תמיד בהישג יד.")
            self.catalog_title.configure(text="מועדפים")
        elif key == "files":
            self.filter_value.set("כל הקבצים")
            self.recipe_filter_value.set("כל הפריטים")
            self.page_title_label.configure(text="כל הקבצים")
            self.page_subtitle_label.configure(text="גם המסמכים, התמונות והקבצים ששמרת בספרייה.")
            self.catalog_title.configure(text="כל הפריטים בספרייה")
        desired_view = "רשימה" if key == "files" else "גלריה"
        self._update_navigation_active()
        if self.view_mode.get() != desired_view:
            self.view_mode.set(desired_view)
            self.change_view_mode()
        else:
            self.refresh()

    def _set_view(self, mode):
        if mode not in ("רשימה", "גלריה") or self.view_mode.get() == mode:
            return
        self.view_mode.set(mode)
        self.change_view_mode()

    def _rebuild_shell(self):
        self._embedded_pages = []
        try:
            self.shell.destroy()
        except Exception:
            pass
        self._build()
        self.refresh()


    def change_view_mode(self):
        """Switch between list and gallery views and remember the choice."""
        mode = self.view_mode.get()
        self.library.set_setting("view_mode", mode)
        self._update_view_buttons()
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
                if hasattr(card, "set_outline"):
                    card.set_outline(self.colors["accent"] if selected else self.colors["border"])
                else:
                    card.configure(
                        highlightbackground=self.colors["accent"] if selected else self.colors["border"],
                        highlightthickness=3 if selected else 1
                    )
            except tk.TclError:
                pass


    def _set_gallery_image_cache_limit(self, image_count: int) -> int:
        """Scale the thumbnail cache to the library, while keeping an estimated RAM budget."""
        try:
            available_images = max(0, int(image_count))
        except (TypeError, ValueError, OverflowError):
            available_images = 0
        limit = min(available_images, GALLERY_CACHE_MAX_IMAGES)
        self._gallery_image_cache_limit = limit
        cache = getattr(self, "_gallery_image_cache", None)
        if cache is not None:
            while len(cache) > limit:
                cache.popitem(last=False)
        return limit

    def _gallery_photo(self, image_path: str):
        """Return a cached gallery thumbnail, refreshing it when the source file changes."""
        if not image_path or Image is None or ImageTk is None or ImageOps is None:
            return None
        cache = getattr(self, "_gallery_image_cache", None)
        if cache is None:
            cache = OrderedDict()
            self._gallery_image_cache = cache
        try:
            path = Path(image_path)
            stat = path.stat()
            normalized_path = os.path.normcase(str(path.resolve()))
            cache_key = (normalized_path, stat.st_mtime_ns, stat.st_size, 320, 145)

            cached = cache.pop(cache_key, None)
            if cached is not None:
                cache[cache_key] = cached
                return cached

            # A modified source must not leave an older thumbnail behind in the cache.
            for old_key in tuple(cache):
                if old_key[0] == normalized_path:
                    cache.pop(old_key, None)

            with Image.open(path) as opened:
                source = ImageOps.exif_transpose(opened).convert("RGB")
                source = ImageOps.fit(source, (320, 145))
                try:
                    from PIL import ImageDraw
                    mask = Image.new("L", source.size, 0)
                    ImageDraw.Draw(mask).rounded_rectangle(
                        (0, 0, source.width - 1, source.height - 1),
                        radius=14, fill=255
                    )
                    source.putalpha(mask)
                except Exception:
                    pass
                preview = ImageTk.PhotoImage(source.copy())

            cache[cache_key] = preview
            cache_limit = max(0, int(getattr(self, "_gallery_image_cache_limit", GALLERY_CACHE_MAX_IMAGES)))
            while len(cache) > cache_limit:
                cache.popitem(last=False)
            return preview
        except Exception:
            return None


    def _render_gallery(self, ranked):
        for child in self.gallery_inner.winfo_children():
            child.destroy()
        self.gallery_photos = []
        self.gallery_cards = {}
        self._gallery_ranked = ranked
        try:
            available_width = max(1, self.gallery_canvas.winfo_width())
        except tk.TclError:
            available_width = 900
        columns = 3 if available_width >= 840 else 2
        for column in range(columns):
            self.gallery_inner.grid_columnconfigure(column, weight=1, uniform="recipe-cards")

        if not ranked:
            empty_shell = RoundedPanel(
                self.gallery_inner, surface=self.colors["row_alt"],
                background=self.colors["surface"], outline=self.colors["border"],
                shadow=self.colors["shadow"], radius=22, padding=2,
                height=285, show_shadow=False
            )
            empty_shell.grid(row=0, column=0, columnspan=columns, sticky="nsew", padx=16, pady=20)
            empty = empty_shell.content
            tk.Label(empty, text="✦", bg=self.colors["row_alt"], fg=self.colors["accent"],
                     font=("Segoe UI Symbol", 38)).pack(pady=(42, 8))
            tk.Label(empty, text="הספרייה שלך מחכה למתכונים הראשונים",
                     bg=self.colors["row_alt"], fg=self.colors["navy"],
                     font=("Segoe UI", 16, "bold")).pack(pady=(4, 5))
            tk.Label(empty, text="הוסיפו מתכון, תמונה או תיקייה שלמה — והכול יופיע כאן.",
                     bg=self.colors["row_alt"], fg=self.colors["muted"],
                     font=("Segoe UI", 10), wraplength=460, justify="center").pack()
            ModernButton(
                empty, text="＋  הוספת מתכונים", command=self.add_files,
                width=170, height=42, bg=self.colors["accent"],
                fg="#FFFFFF", hover_bg=self.colors["accent_hover"],
                active_bg=self.colors["accent_hover"], radius=14,
                font=("Segoe UI", 10, "bold"), canvas_bg=self.colors["row_alt"]
            ).pack(pady=(17, 36))
            self.gallery_canvas.configure(scrollregion=self.gallery_canvas.bbox("all"))
            return

        gallery_limit = max(48, int(getattr(self, "_gallery_limit", 60)))
        visible_rows = ranked[:gallery_limit]
        for index, (_score, row, file_size, _imported) in enumerate(visible_rows):
            item_id = str(row["id"])
            selected = item_id == str(self.selected_item_id)
            card = RoundedPanel(
                self.gallery_inner, surface=self.colors["surface"],
                background=self.colors["surface"],
                outline=self.colors["accent"] if selected else self.colors["border"],
                shadow=self.colors["shadow"], radius=18, padding=1,
                width=280, height=264, show_shadow=True
            )
            card.grid(
                row=index // columns, column=columns - 1 - (index % columns),
                sticky="nsew", padx=9, pady=9
            )
            self.gallery_cards[item_id] = card
            body = card.content
            data = _recipe_data_from_row(row)
            image_path = str(data.get("image_path") or "")
            if not image_path and row["extension"].lower() in IMAGE_TYPES:
                image_path = row["stored_path"]

            media = tk.Frame(body, bg=self.colors["cover"], height=145)
            media.pack(fill="x", padx=10, pady=(10, 8))
            media.pack_propagate(False)
            preview = self._gallery_photo(image_path) if image_path else None
            if preview is not None:
                self.gallery_photos.append(preview)

            if preview is not None:
                image_label = tk.Label(media, image=preview, bg=self.colors["cover"], cursor="hand2")
                image_label.image = preview
                image_label.pack(fill="both", expand=True)
            else:
                placeholder = tk.Frame(media, bg=self.colors["cover"])
                placeholder.pack(fill="both", expand=True)
                tk.Label(
                    placeholder, text="✦", bg=self.colors["cover"],
                    fg=self.colors["accent"], font=("Segoe UI Symbol", 32)
                ).pack(expand=True)
                tk.Label(
                    placeholder, text="מתכון מהספרייה", bg=self.colors["cover"],
                    fg=self.colors["muted"], font=("Segoe UI", 8)
                ).pack(pady=(0, 12))

            if row["is_favorite"]:
                tk.Label(
                    media, text="★", bg=self.colors["surface"], fg=self.colors["accent"],
                    font=("Segoe UI", 11, "bold"), padx=6, pady=2
                ).place(relx=0.96, rely=0.08, anchor="ne")

            tk.Label(
                body, text=row["title"], bg=self.colors["surface"],
                fg=self.colors["navy"], font=("Segoe UI", 11, "bold"),
                justify="right", anchor="e", wraplength=245, height=2
            ).pack(fill="x", padx=14, pady=(2, 1))

            prep = str(data.get("prep_minutes") or "").strip()
            cook = str(data.get("cook_minutes") or "").strip()
            servings = str(data.get("servings") or "").strip()
            meta_parts = []
            if prep or cook:
                total = sum(int(value) for value in (prep, cook) if value.isdigit())
                meta_parts.append(f"{total} דק׳")
            if servings:
                meta_parts.append(f"{servings} מנות")
            if not meta_parts and data.get("category"):
                meta_parts.append(str(data.get("category")).split("/")[-1])
            rating = str(data.get("rating") or "").strip()
            if rating.isdigit() and int(rating) > 0:
                meta_parts.append("★ " + rating)
            tk.Label(
                body, text="   ·   ".join(meta_parts) if meta_parts else "מתכון",
                bg=self.colors["surface"], fg=self.colors["muted"],
                font=("Segoe UI", 8), justify="right", anchor="e"
            ).pack(fill="x", padx=14, pady=(1, 10))
            self._bind_gallery_click(card, item_id)

        if len(ranked) > len(visible_rows):
            more_row = (len(visible_rows) + columns - 1) // columns
            more_button = ModernButton(
                self.gallery_inner,
                text=f"הצגת עוד {min(48, len(ranked) - len(visible_rows))} מתכונים  ↓",
                command=self._load_more_recipes,
                width=220, height=42,
                bg=self.colors["pill"], fg=self.colors["pill_text"],
                hover_bg=self.colors["soft"], active_bg=self.colors["sidebar_active"],
                radius=14, font=("Segoe UI", 9, "bold"),
                canvas_bg=self.colors["surface"]
            )
            more_button.grid(
                row=more_row, column=0, columnspan=columns,
                padx=12, pady=(10, 20)
            )
        self.gallery_canvas.configure(scrollregion=self.gallery_canvas.bbox("all"))

    def _load_more_recipes(self):
        self._gallery_limit = max(48, int(getattr(self, "_gallery_limit", 60))) + 48
        self._render_gallery(getattr(self, "_gallery_ranked", []))


    def _update_filter_options(self):
        options = {
            "כל הקבצים": ("all", None),
            "★ מועדפים": ("favorites", None),
        }
        for collection in self.library.collections():
            options[f"📁 {collection['name']}"] = ("collection", int(collection["id"]))
        current = self.filter_value.get()
        self._filter_options = options
        if getattr(self, "filter_combo", None) is not None:
            self.filter_combo.configure(values=list(options))
        self.filter_value.set(current if current in options else "כל הקבצים")
        categories = ["כל הקטגוריות"] + self.library.all_categories()
        tags = ["כל התגיות"] + self.library.all_tags()
        self.category_combo.configure(values=categories)
        if getattr(self, "tag_combo", None) is not None:
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
        folder_filter = self.folder_filter_value.get()
        recipe_filter = self.recipe_filter_value.get()
        sort_mode = self.sort_value.get()
        need_file_size = self.view_mode.get() == "רשימה" or sort_mode == "גודל: גדול לקטן"
        tags_by_item = (
            self.library.tags_by_item(row["id"] for row in rows)
            if tag_filter != "כל התגיות" or query else {}
        )
        recipe_by_id = {}
        ranked = []
        gallery_image_count = 0
        for row in rows:
            kind = file_kind(row["extension"])
            if type_filter != "כל הסוגים" and kind != type_filter:
                continue
            recipe = _recipe_data_from_row(row)
            item_id = int(row["id"])
            recipe_by_id[item_id] = recipe
            if str(recipe.get("image_path") or "").strip() or str(row["extension"] or "").lower() in IMAGE_TYPES:
                gallery_image_count += 1
            category = str(recipe.get("category") or "")
            folder_path = str(recipe.get("source_folder") or "").strip("/")
            if folder_filter != "כל התיקיות" and not (
                folder_path == folder_filter or folder_path.startswith(folder_filter + "/")
            ):
                continue
            is_recipe = is_recipe_record(recipe, row["content"] or "")
            if recipe_filter == "מתכונים בלבד" and not is_recipe:
                continue
            if recipe_filter == "מסמכים וקבצים" and is_recipe:
                continue
            if recipe_filter == "מתכונים חסרים" and (not is_recipe or not recipe_missing_parts(recipe)):
                continue
            if category_filter != "כל הקטגוריות" and category != category_filter:
                continue
            item_tags = tags_by_item.get(item_id, [])
            if tag_filter != "כל התגיות" and tag_filter not in item_tags:
                continue
            imported = self._parse_imported_date(row["imported_at"])
            if date_from and (imported is None or imported.date() < date_from):
                continue
            if date_to and (imported is None or imported.date() > date_to):
                continue
            score = 1 if not query else relevance(
                row["title"],
                recipe_search_text(row["title"], row["content"], recipe, item_tags),
                query,
            )
            if score > 0:
                if need_file_size:
                    try:
                        file_size = Path(row["stored_path"]).stat().st_size
                    except OSError:
                        file_size = 0
                else:
                    file_size = 0
                ranked.append((score, row, file_size, imported))

        self._set_gallery_image_cache_limit(gallery_image_count)

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
                key=lambda x: int(recipe_by_id.get(int(x[1]["id"]), {}).get("rating") or 0)
                if str(recipe_by_id.get(int(x[1]["id"]), {}).get("rating") or "0").isdigit() else 0,
                reverse=True
            )
        elif sort_mode == "נצפו לאחרונה":
            ranked.sort(key=lambda x: x[1]["last_opened"] or "", reverse=True)
        elif query:
            ranked.sort(key=lambda x: (x[0], x[1]["id"]), reverse=True)
        else:
            ranked.sort(key=lambda x: (x[3] or datetime.min, x[1]["id"]), reverse=True)

        ranked_ids = {str(entry[1]["id"]) for entry in ranked}
        if (
            ranked
            and self.current_page in ("recipes", "favorites")
            and str(selected_id or "") not in ranked_ids
        ):
            selected_id = str(ranked[0][1]["id"])

        existing_tree_items = self.tree.get_children()
        if existing_tree_items:
            self.tree.delete(*existing_tree_items)
        self.current_rows = {str(entry[1]["id"]): entry[1] for entry in ranked}
        if self.view_mode.get() == "רשימה":
            for index, (_score, row, file_size, _imported) in enumerate(ranked):
                iid = str(row["id"])
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

        if self.view_mode.get() == "גלריה":
            self._render_gallery(ranked)
        else:
            for child in self.gallery_inner.winfo_children():
                child.destroy()
            self.gallery_photos = []
            self.gallery_cards = {}
            self._gallery_ranked = ranked
        if hasattr(self, "count_label"):
            self.count_label.configure(text=f"{len(ranked)} מתכונים")
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
            self._render_reader(None)
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
            self.favorite_button.set_text("☆  מועדפים")
            self.favorite_button.set_disabled(True)
        elif row["is_favorite"]:
            self.favorite_button.set_text("★  במועדפים")
            self.favorite_button.set_disabled(False)
        else:
            self.favorite_button.set_text("☆  הוספה למועדפים")
            self.favorite_button.set_disabled(False)

    def show_selected(self, _event=None):
        row = self.selected_row()
        if not row:
            self._update_selection_actions(None)
            self._render_reader(None)
            return
        self.selected_item_id = str(row["id"])
        self._update_selection_actions(row)
        self._render_reader(row)
        self._render_gallery_cards_selection()

    def _render_reader(self, row):
        """Render a premium, easy-to-read recipe card inside the home screen."""
        if not hasattr(self, "reader_inner"):
            return
        for child in self.reader_inner.winfo_children():
            child.destroy()
        self.favorite_button = None
        self.reader_photos = []
        c = self.colors
        if row is None:
            tk.Label(self.reader_inner, text="✦", bg=c["surface"], fg=c["accent"],
                     font=("Segoe UI Symbol", 42)).pack(fill="x", pady=(52, 8))
            tk.Label(self.reader_inner, text="המתכון הבא שלך מתחיל כאן",
                     bg=c["surface"], fg=c["navy"], font=("Segoe UI", 14, "bold"),
                     wraplength=280, justify="center").pack(fill="x", padx=18)
            tk.Label(self.reader_inner,
                     text="בחרו כרטיס מהספרייה כדי לקרוא את המצרכים ואת שלבי ההכנה.",
                     bg=c["surface"], fg=c["muted"], font=("Segoe UI", 9),
                     wraplength=270, justify="center").pack(fill="x", padx=18, pady=(7, 0))
            return

        data = self.library.get_recipe_data(row["id"])
        image_path = str(data.get("image_path") or "")
        if not image_path and row["extension"].lower() in IMAGE_TYPES:
            image_path = row["stored_path"]
        mode = self.reader_preview_mode.get()
        if mode == "text":
            source_text = str(row["content"] or "").strip()
            if not source_text:
                source_text = "\n\n".join(
                    part for part in (
                        data.get("ingredients", ""),
                        data.get("instructions", ""),
                        data.get("notes", "")
                    ) if part
                )
            if not source_text:
                source_text = "לא נשמר טקסט שחולץ מהקובץ הזה."
            text_shell = RoundedPanel(
                self.reader_inner, surface=c["row_alt"], background=c["surface"],
                outline=c["border"], shadow=c["shadow"], radius=16,
                padding=2, height=245, show_shadow=False
            )
            text_shell.pack(fill="x", padx=12, pady=(10, 11))
            text_body = tk.Frame(text_shell.content, bg=c["row_alt"])
            text_body.pack(fill="both", expand=True, padx=8, pady=8)
            text_area = tk.Text(
                text_body, height=10, wrap="word", font=("Segoe UI", 10),
                bg=c["row_alt"], fg=c["text"], relief="flat", bd=0,
                padx=9, pady=7
            )
            text_scroll = ttk.Scrollbar(text_body, orient="vertical", command=text_area.yview)
            text_area.configure(yscrollcommand=text_scroll.set)
            text_scroll.pack(side="left", fill="y")
            text_area.pack(side="right", fill="both", expand=True)
            text_area.insert("1.0", source_text)
            text_area.tag_configure("rtl", justify="right")
            text_area.tag_add("rtl", "1.0", "end")
            text_area.configure(state="disabled")
            photo = None
        else:
            photo = None
        if mode == "image" and image_path and Image is not None and ImageTk is not None:
            try:
                with Image.open(image_path) as source:
                    source = ImageOps.exif_transpose(source).convert("RGB")
                    source.thumbnail((304, 430), getattr(getattr(Image, "Resampling", Image), "LANCZOS", 1))
                    try:
                        from PIL import ImageDraw
                        mask = Image.new("L", source.size, 0)
                        ImageDraw.Draw(mask).rounded_rectangle(
                            (0, 0, source.width - 1, source.height - 1), radius=16, fill=255
                        )
                        source.putalpha(mask)
                    except Exception:
                        pass
                    photo = ImageTk.PhotoImage(source.copy())
                self.reader_photos.append(photo)
            except Exception:
                photo = None
        if mode == "image" and photo is None and row["extension"].lower() == ".pdf" and fitz is not None and Image is not None and ImageTk is not None:
            try:
                with fitz.open(row["stored_path"]) as document:
                    if document.page_count:
                        pix = document[0].get_pixmap(matrix=fitz.Matrix(1.0, 1.0), alpha=False)
                        source = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                        source.thumbnail((304, 430), getattr(getattr(Image, "Resampling", Image), "LANCZOS", 1))
                        photo = ImageTk.PhotoImage(source.copy())
                        self.reader_photos.append(photo)
            except Exception:
                photo = None

        if mode == "image" and photo is not None:
            image_label = tk.Label(self.reader_inner, image=photo, bg=c["surface"], highlightthickness=0, bd=0)
            image_label.image = photo
            image_label.pack(anchor="center", padx=12, pady=(10, 11))
        elif mode == "image":
            cover = RoundedPanel(
                self.reader_inner, surface=c["cover"], background=c["surface"],
                outline=c["border"], shadow=c["shadow"], radius=17,
                padding=1, height=124, show_shadow=False
            )
            cover.pack(fill="x", padx=12, pady=(10, 11))
            tk.Label(cover.content, text="✦", bg=c["cover"], fg=c["accent"],
                     font=("Segoe UI Symbol", 34)).pack(expand=True)
            tk.Label(cover.content, text="מתכון מהספרייה שלך", bg=c["cover"],
                     fg=c["muted"], font=("Segoe UI", 8)).pack(pady=(0, 12))

        tk.Label(
            self.reader_inner, text=row["title"], bg=c["surface"], fg=c["navy"],
            font=("Segoe UI", 17, "bold"), justify="right", anchor="e", wraplength=278
        ).pack(fill="x", padx=15, pady=(0, 5))

        meta = []
        if data.get("category"):
            meta.append(str(data["category"]).split("/")[-1])
        if data.get("prep_minutes"):
            meta.append(f"הכנה {data['prep_minutes']} דק׳")
        if data.get("cook_minutes"):
            meta.append(f"בישול {data['cook_minutes']} דק׳")
        if data.get("servings"):
            meta.append(f"{data['servings']} מנות")
        if data.get("rating"):
            meta.append("★ " + str(data["rating"]))
        tk.Label(
            self.reader_inner, text="   ·   ".join(meta) if meta else "מהאוסף האישי שלך",
            bg=c["surface"], fg=c["muted"], font=("Segoe UI", 8),
            justify="right", anchor="e", wraplength=280
        ).pack(fill="x", padx=15, pady=(0, 11))

        actions = tk.Frame(self.reader_inner, bg=c["surface"])
        actions.pack(fill="x", padx=10, pady=(0, 8))
        self.favorite_button = ModernButton(
            actions, text="★  מועדף" if row["is_favorite"] else "☆  הוספה למועדפים",
            command=self.toggle_favorite, width=145, height=36,
            bg=c["pill"], fg=c["pill_text"], hover_bg=c["soft"],
            active_bg=c["sidebar_active"], active_fg="#FFFFFF",
            radius=12, font=("Segoe UI", 8, "bold"), canvas_bg=c["surface"]
        )
        self.favorite_button.pack(side="right", fill="x", expand=True, padx=3)
        ModernButton(
            actions, text="עריכה", command=self.edit_recipe,
            width=80, height=36, bg=c["accent"], fg="#FFFFFF",
            hover_bg=c["accent_hover"], active_bg=c["accent_hover"],
            radius=12, font=("Segoe UI", 9, "bold"), canvas_bg=c["surface"]
        ).pack(side="right", fill="x", expand=True, padx=3)

        tools = tk.Frame(self.reader_inner, bg=c["surface"])
        tools.pack(fill="x", padx=10, pady=(0, 12))
        ModernButton(
            tools, text="מצב בישול", command=lambda: self.open_recipe_center(initial_item_id=int(row["id"])),
            width=130, height=34, bg=c["row_alt"], fg=c["navy"],
            hover_bg=c["soft"], radius=12, font=("Segoe UI", 8, "bold"),
            canvas_bg=c["surface"]
        ).pack(side="right", fill="x", expand=True, padx=3)
        ModernButton(
            tools, text="הדפסה", command=self.print_selected_recipe,
            width=90, height=34, bg=c["row_alt"], fg=c["navy"],
            hover_bg=c["soft"], radius=12, font=("Segoe UI", 8, "bold"),
            canvas_bg=c["surface"]
        ).pack(side="right", fill="x", expand=True, padx=3)

        def add_section(heading, content, min_lines=4):
            section = tk.Frame(self.reader_inner, bg=c["surface"])
            section.pack(fill="x", padx=13, pady=(4, 8))
            tk.Label(section, text=heading, bg=c["surface"], fg=c["navy"],
                     font=("Segoe UI", 10, "bold"), anchor="e").pack(fill="x", pady=(0, 5))
            text_value = content or "לא נשמר עדיין מידע בחלק הזה."
            visual_lines = sum(max(1, (len(line) + 31) // 32) for line in str(text_value).splitlines()) or 1
            box_height = max(min_lines, min(18, visual_lines + 1))
            text_shell = RoundedPanel(
                section, surface=c["row_alt"], background=c["surface"],
                outline=c["border"], shadow=c["shadow"],
                radius=13, padding=1, height=max(74, 21 * box_height), show_shadow=False
            )
            text_shell.pack(fill="x")
            area = tk.Text(
                text_shell.content, height=box_height, wrap="word",
                font=("Segoe UI", 10), bg=c["row_alt"], fg=c["text"],
                relief="flat", bd=0, padx=11, pady=9,
                spacing1=2, spacing3=3
            )
            area.pack(fill="both", expand=True, padx=2, pady=2)
            area.insert("1.0", text_value)
            area.tag_configure("rtl", justify="right")
            area.tag_add("rtl", "1.0", "end")
            area.configure(state="disabled")

        content = row["content"] or ""
        parsed = parse_recipe_text(content)
        ingredients = data.get("ingredients") or parsed.get("ingredients", "")
        instructions = data.get("instructions") or parsed.get("instructions", "")
        add_section("מצרכים", ingredients, 4)
        add_section("אופן ההכנה", instructions, 5)
        if data.get("notes"):
            add_section("הערות אישיות", data["notes"], 3)
        if not ingredients and not instructions and content:
            add_section("הטקסט המקורי", content[:5000], 7)
        self.reader_canvas.update_idletasks()
        self.reader_canvas.configure(scrollregion=self.reader_canvas.bbox("all"))
        self.reader_canvas.yview_moveto(0)


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
        dialog = self._push_embedded_page("ניהול אוספים")
        self._embedded_header(dialog, "ניהול אוספים", "סידור המתכונים באוספים אישיים.")
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
        ttk.Button(buttons, text="חזרה לספרייה", command=lambda: self._pop_embedded_page(dialog)).pack(side="left")

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

        window = self._push_embedded_page(f"פרטי מתכון – {row['title']}")
        self._embedded_header(window, "פרטי המתכון", row["title"])
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
            result = self._push_embedded_page(f"כמויות ל־{target} מנות")
            self._embedded_header(result, f"כמויות ל־{target} מנות", "חישוב כמויות לפי מספר המנות שבחרתם.")
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
            result = self._push_embedded_page("השוואת גרסאות")
            self._embedded_header(result, "השוואת גרסאות", "השוואה בין הטקסט השמור לבין הגרסה הנוכחית.")
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
                self._pop_embedded_page(window)
            except Exception as exc:
                messagebox.showerror("שמירת מתכון", str(exc), parent=window)

        ttk.Button(buttons, text="שמירת הפרטים", style="Primary.TButton", command=save_recipe).pack(side="right")
        ttk.Button(buttons, text="חזרה", command=lambda: self._pop_embedded_page(window)).pack(side="left")


    def open_ocr_scanner(self):
        """Open the OCR workspace as a page inside the main application."""
        window = self._push_embedded_page("סורק מתכונים חכם")
        self._embedded_header(
            window, "סורק מתכונים חכם",
            "סריקה, תיקון הזיהוי ושמירת המתכון — בלי לפתוח חלון נוסף."
        )
        outer = tk.Frame(window, bg=self.colors["background"], padx=18, pady=8)
        outer.pack(fill="both", expand=True)

        state = {"paths": [], "result": None, "current_path": None, "preview": None}
        toolbar = tk.Frame(outer, bg=self.colors["surface"], padx=12, pady=10)
        toolbar.pack(fill="x", pady=(0, 10))
        list_area = tk.Frame(outer, bg=self.colors["surface"], padx=10, pady=8)
        list_area.pack(fill="x", pady=(0, 10))
        tk.Label(list_area, text="קבצים לסריקה", bg=self.colors["surface"],
                 fg=self.colors["navy"], font=("Segoe UI", 10, "bold"), anchor="e").pack(fill="x")
        file_list = tk.Listbox(
            list_area, height=3, selectmode="browse", activestyle="none",
            font=("Segoe UI", 10), bg=self.colors["row_alt"], fg=self.colors["text"],
            relief="flat", highlightthickness=0, justify="right"
        )
        file_list.pack(fill="x", pady=(6, 0))

        body = tk.Frame(outer, bg=self.colors["background"])
        body.pack(fill="both", expand=True)
        source_shell = RoundedPanel(
            body, surface=self.colors["surface"], background=self.colors["background"],
            outline=self.colors["border"], shadow=self.colors["shadow"],
            radius=19, padding=2, width=440, show_shadow=True
        )
        source_shell.pack(side="left", fill="both", padx=(0, 12))
        source_panel = source_shell.content
        tk.Label(source_panel, text="המסמך המקורי", bg=self.colors["surface"],
                 fg=self.colors["navy"], font=("Segoe UI", 12, "bold"), anchor="e").pack(fill="x", padx=12, pady=10)
        source_view = tk.Label(
            source_panel, text="בחרו קובץ כדי להציג אותו כאן",
            bg=self.colors["row_alt"], fg=self.colors["muted"],
            font=("Segoe UI", 11), justify="center"
        )
        source_view.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        editor_shell = RoundedPanel(
            body, surface=self.colors["surface"], background=self.colors["background"],
            outline=self.colors["border"], shadow=self.colors["shadow"],
            radius=19, padding=2, show_shadow=True
        )
        editor_shell.pack(side="left", fill="both", expand=True)
        editor_inner = tk.Frame(editor_shell.content, bg=self.colors["surface"], padx=14, pady=12)
        editor_inner.pack(fill="both", expand=True)
        title_var = tk.StringVar()
        tk.Label(editor_inner, text="שם המתכון", bg=self.colors["surface"],
                 fg=self.colors["navy"], font=("Segoe UI", 10, "bold"), anchor="e").pack(fill="x")
        ttk.Entry(editor_inner, textvariable=title_var, justify="right").pack(fill="x", pady=(4, 10), ipady=3)

        def make_editor(label, height):
            box = tk.Frame(editor_inner, bg=self.colors["surface"])
            box.pack(fill="both", expand=True, pady=(0, 8))
            tk.Label(box, text=label, bg=self.colors["surface"], fg=self.colors["navy"],
                     font=("Segoe UI", 10, "bold"), anchor="e").pack(fill="x", pady=(0, 4))
            area = tk.Text(
                box, height=height, wrap="word", font=("Segoe UI", 10),
                bg=self.colors["row_alt"], fg=self.colors["text"], relief="flat",
                padx=10, pady=8, spacing1=2, spacing3=3
            )
            area.pack(fill="both", expand=True)
            area.tag_configure("rtl", justify="right")
            return area

        ingredients_box = make_editor("מצרכים", 7)
        instructions_box = make_editor("אופן ההכנה", 8)
        raw_box = make_editor("טקסט מלא שזוהה – ניתן להשוות למקור", 5)
        status_var = tk.StringVar(value="בחרו קובץ תמונה או PDF כדי להתחיל.")
        tk.Label(editor_inner, textvariable=status_var, bg=self.colors["surface"],
                 fg=self.colors["muted"], font=("Segoe UI", 9), anchor="e").pack(fill="x", pady=(0, 6))

        def add_files():
            selected = filedialog.askopenfilenames(
                parent=window, title="בחירת תמונות או מסמכים לסריקה",
                filetypes=[
                    ("תמונות ו-PDF", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp *.pdf"),
                    ("כל הקבצים", "*.*")
                ]
            )
            if not selected:
                return
            for value in selected:
                path = Path(value)
                if path not in state["paths"]:
                    state["paths"].append(path)
                    file_list.insert("end", path.name)
            if file_list.size() and not file_list.curselection():
                file_list.selection_set(0)

        def remove_file():
            selection = file_list.curselection()
            if not selection:
                return
            index = selection[0]
            file_list.delete(index)
            state["paths"].pop(index)
            state["result"] = None
            state["current_path"] = None
            state["preview"] = None
            source_view.configure(image="", text="בחרו קובץ כדי להציג אותו כאן")
            source_view.image = None

        def set_box(box, value):
            box.configure(state="normal")
            box.delete("1.0", "end")
            box.insert("1.0", value or "")
            box.tag_add("rtl", "1.0", "end")

        def scan_failed(error):
            scan_button.configure(state="normal")
            save_button.configure(state="normal")
            status_var.set("הסריקה לא הושלמה.")
            messagebox.showerror("שגיאה בסריקה", error, parent=window)

        def finish_scan(payload):
            state["result"] = payload
            state["current_path"] = payload["path"]
            state["preview"] = payload["preview"]
            title_var.set(payload["parsed"].get("title") or payload["path"].stem)
            set_box(ingredients_box, payload["parsed"].get("ingredients", ""))
            set_box(instructions_box, payload["parsed"].get("instructions", ""))
            set_box(raw_box, payload["text"])
            preview = payload["preview"]
            if preview is not None and ImageTk is not None:
                preview.thumbnail((400, 700), getattr(getattr(Image, "Resampling", Image), "LANCZOS", 1))
                photo = ImageTk.PhotoImage(preview.copy())
                source_view.configure(image=photo, text="")
                source_view.image = photo
            else:
                source_view.configure(image="", text=payload["path"].name)
                source_view.image = None
            confidence = payload["confidence"]
            confidence_text = f"רמת ביטחון ממוצעת: {confidence:.0f}%" if confidence else "לא התקבל ציון ביטחון"
            status_var.set(f"הסריקה הסתיימה. {confidence_text} · בדקו את השדות לפני השמירה.")
            scan_button.configure(state="normal")
            save_button.configure(state="normal")

        def scan_selected():
            selection = file_list.curselection()
            if not selection:
                messagebox.showinfo("סריקת מתכון", "בחרו קובץ מהרשימה.", parent=window)
                return
            path = state["paths"][selection[0]]
            if pytesseract is None or Image is None:
                messagebox.showerror("OCR אינו זמין", "מנוע זיהוי הטקסט אינו זמין בחבילה.", parent=window)
                return
            state["result"] = None
            scan_button.configure(state="disabled")
            save_button.configure(state="disabled")
            status_var.set("מנתח את המסמך ומחפש את הכותרת הבולטת בעמוד…")

            def worker():
                try:
                    text_parts = []
                    title_candidate = ""
                    preview_image = None
                    confidence_values = []
                    if path.suffix.lower() == ".pdf":
                        if fitz is None:
                            raise RuntimeError("קריאת PDF אינה זמינה בחבילה.")
                        with fitz.open(path) as document:
                            for page_index, page in enumerate(document):
                                page_text = page.get_text("text").strip()
                                page_image = None
                                result = {"text": "", "title": "", "confidence": 0}
                                if not page_text or page_index == 0:
                                    pix = page.get_pixmap(matrix=fitz.Matrix(2.1, 2.1), alpha=False)
                                    import io
                                    with Image.open(io.BytesIO(pix.tobytes("png"))) as rendered:
                                        page_image = rendered.convert("RGB")
                                    if page_index == 0:
                                        preview_image = page_image.copy()
                                        title_candidate = best_title_from_pdf_page(page) or ""
                                    result = recognize_image(page_image, pytesseract)
                                if not page_text:
                                    page_text = result.get("text", "")
                                if page_index == 0 and result.get("title"):
                                    title_candidate = result["title"] or title_candidate
                                if result.get("confidence"):
                                    confidence_values.append(result["confidence"])
                                text_parts.append(page_text)
                    elif path.suffix.lower() in IMAGE_TYPES:
                        with Image.open(path) as source:
                            preview_image = ImageOps.exif_transpose(source).convert("RGB")
                            result = recognize_image(preview_image, pytesseract)
                        title_candidate = result.get("title", "")
                        text_parts.append(result.get("text", ""))
                        if result.get("confidence"):
                            confidence_values.append(result["confidence"])
                    else:
                        raise ValueError("הקובץ אינו תמונה או PDF נתמך.")
                    full_text = "\n\n".join(part for part in text_parts if part).strip()
                    if not full_text:
                        raise RuntimeError("לא זוהה טקסט. נסו תמונה חדה יותר או סריקה באיכות גבוהה.")
                    parsed = parse_recipe_text(full_text)
                    if title_candidate:
                        parsed["title"] = title_candidate
                    payload = {
                        "path": path, "text": full_text, "parsed": parsed,
                        "title": title_candidate, "preview": preview_image,
                        "confidence": round(sum(confidence_values) / len(confidence_values), 1) if confidence_values else 0,
                    }
                    window.after(0, lambda value=payload: finish_scan(value))
                except Exception as exc:
                    error = str(exc)
                    try:
                        window.after(0, lambda value=error: scan_failed(value))
                    except Exception:
                        pass

            self._scan_thread = threading.Thread(target=worker, daemon=True)
            self._scan_thread.start()

        def save_recipe():
            result = state.get("result")
            if not result:
                messagebox.showinfo("שמירת מתכון", "יש לסרוק מסמך לפני השמירה.", parent=window)
                return
            path = Path(result["path"])
            title = safe_title(title_var.get().strip())
            ingredients = ingredients_box.get("1.0", "end-1c").strip()
            instructions = instructions_box.get("1.0", "end-1c").strip()
            raw_text = raw_box.get("1.0", "end-1c").strip()
            if not title or (not ingredients and not instructions):
                messagebox.showwarning("שמירת מתכון", "יש להזין שם מתכון ולפחות מצרכים או הוראות הכנה.", parent=window)
                return
            digest = file_digest(path)
            existing = self.library.db.execute("SELECT * FROM items WHERE digest=?", (digest,)).fetchone()
            if existing:
                item_id = int(existing["id"])
                row = existing
            else:
                OCR_FILE_CACHE[str(path.resolve())] = {"text": raw_text, "title": title}
                added, message = self.library.import_one(path)
                if not added:
                    messagebox.showerror("שמירת מתכון", message, parent=window)
                    return
                row = self.library.db.execute("SELECT * FROM items WHERE digest=?", (digest,)).fetchone()
                if not row:
                    messagebox.showerror("שמירת מתכון", "הקובץ נשמר אך לא נמצאה הרשומה שלו.", parent=window)
                    return
                item_id = int(row["id"])
            recipe = self.library.get_recipe_data(item_id)
            recipe.update({
                "ingredients": ingredients,
                "instructions": instructions,
                "is_recipe": True,
            })
            self.library.db.execute("UPDATE items SET content=? WHERE id=?", (raw_text, item_id))
            self.library.db.commit()
            if title != row["title"]:
                self.library.rename(item_id, title)
            stored_row = self.library.db.execute(
                "SELECT stored_path FROM items WHERE id=?", (item_id,)
            ).fetchone()
            if path.suffix.lower() in IMAGE_TYPES and stored_row:
                recipe["image_path"] = str(stored_row["stored_path"])
            self.library.set_recipe_data(item_id, recipe)
            self.current_page = "recipes"
            self.selected_item_id = str(item_id)
            self.filter_value.set("כל הקבצים")
            self.recipe_filter_value.set("מתכונים בלבד")
            self.view_mode.set("גלריה")
            self._pop_embedded_page(window)
            self._rebuild_shell()
            self.status_value.set("המתכון נסרק ונשמר בספרייה.")
            messagebox.showinfo("המתכון נשמר", "המתכון נוסף לספרייה בהצלחה.", parent=self.root)

        tk.Button(toolbar, text="＋ בחירת קבצים", command=add_files,
                  bg=self.colors["accent"], fg="#FFFFFF", activebackground=self.colors["accent_hover"],
                  relief="flat", bd=0, padx=16, pady=8, cursor="hand2",
                  font=("Segoe UI", 9, "bold")).pack(side="right", padx=4)
        tk.Button(toolbar, text="הסרת קובץ", command=remove_file,
                  bg=self.colors["row_alt"], fg=self.colors["navy"], activebackground=self.colors["selected"],
                  relief="flat", bd=0, padx=14, pady=8, cursor="hand2",
                  font=("Segoe UI", 9)).pack(side="right", padx=4)
        scan_button = tk.Button(toolbar, text="⌕ סריקה וזיהוי", command=scan_selected,
                                bg=self.colors["sidebar"], fg="#FFFFFF", activebackground=self.colors["sidebar_active"],
                                relief="flat", bd=0, padx=18, pady=8, cursor="hand2",
                                font=("Segoe UI", 9, "bold"))
        scan_button.pack(side="left", padx=4)
        save_button = tk.Button(toolbar, text="שמירת מתכון בספרייה", command=save_recipe,
                                bg=self.colors["accent"], fg="#FFFFFF", activebackground=self.colors["accent_hover"],
                                relief="flat", bd=0, padx=18, pady=8, cursor="hand2",
                                font=("Segoe UI", 9, "bold"))
        save_button.pack(side="left", padx=4)
        tk.Button(toolbar, text="חזרה לספרייה", command=lambda: self._pop_embedded_page(window),
                  bg=self.colors["row_alt"], fg=self.colors["navy"], activebackground=self.colors["selected"],
                  relief="flat", bd=0, padx=14, pady=8, cursor="hand2",
                  font=("Segoe UI", 9)).pack(side="left", padx=4)
        file_list.bind("<<ListboxSelect>>", lambda _event: status_var.set("לחצו על סריקה וזיהוי כדי לנתח את הקובץ הנבחר."))

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

        dialog = self._push_embedded_page("איתור פריטים דומים")
        self._embedded_header(dialog, "איתור מתכונים וקבצים דומים", "הצגת התאמות אפשריות בלי למחוק שום דבר אוטומטית.")
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
        ttk.Button(frame, text="חזרה לספרייה", command=lambda: self._pop_embedded_page(dialog)).pack(anchor="w", pady=(10, 0))

    def open_recycle_bin(self):
        dialog = self._push_embedded_page("סל המחזור")
        self._embedded_header(dialog, "סל המחזור", "שחזור פריטים או מחיקה סופית.")
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
            self._pop_embedded_page(dialog)
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
        ttk.Button(controls, text="חזרה לספרייה", command=lambda: self._pop_embedded_page(dialog)).pack(side="left")

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
        """Import folders without blocking Tk; OCR and hashing run in one background worker."""
        paths = [Path(path) for path in paths]
        if not paths:
            return

        home = self.library.home.resolve()
        known_digests = {
            str(row["digest"]) for row in self.library.db.execute("SELECT digest FROM items").fetchall()
        }
        results: queue.Queue = queue.Queue(maxsize=1)
        cancel_event = threading.Event()
        self._import_cancel_event = cancel_event
        state = {"added": 0, "skipped": 0, "failed": 0, "processed": 0}

        cancel_action = {"callback": lambda: cancel_event.set()}
        progress_window = self._push_embedded_page("ייבוא ספרייה")
        self._embedded_header(
            progress_window, "ייבוא ספרייה", "סריקת הקבצים והכנתם לספרייה.",
            back_command=lambda: cancel_action["callback"](),
            back_text="ביטול ייבוא"
        )
        frame = ttk.Frame(progress_window, padding=18, style="App.TFrame")
        frame.pack(fill="both", expand=True)
        progress_text = tk.StringVar(value=f"מכין לייבוא {len(paths)} קבצים…")
        ttk.Label(frame, textvariable=progress_text, style="Panel.TLabel", anchor="e").pack(fill="x", pady=(0, 10))
        progress = ttk.Progressbar(frame, mode="determinate", maximum=max(1, len(paths)), value=0)
        progress.pack(fill="x", pady=(0, 12))
        counter = tk.StringVar(value=f"0 מתוך {len(paths)}")
        ttk.Label(frame, textvariable=counter, style="Panel.TLabel", anchor="e").pack(fill="x")
        buttons = ttk.Frame(frame, style="App.TFrame")
        buttons.pack(fill="x", pady=(12, 0))
        cancel_button = ttk.Button(buttons, text="ביטול ייבוא")
        cancel_button.pack(side="left")

        def request_cancel():
            if not cancel_event.is_set():
                cancel_event.set()
                cancel_button.configure(state="disabled")
                progress_text.set("מבטל… הקובץ הנוכחי יסיים עיבוד לפני העצירה.")
        cancel_button.configure(command=request_cancel)
        cancel_action["callback"] = request_cancel

        def publish_result(record, allow_after_cancel=False):
            while True:
                try:
                    results.put(record, timeout=0.2)
                    return
                except queue.Full:
                    if cancel_event.is_set() and not allow_after_cancel:
                        return

        def import_worker():
            try:
                for index, original_path in enumerate(paths, 1):
                    if cancel_event.is_set():
                        break
                    try:
                        path = original_path.resolve()
                        if path == home or home in path.parents:
                            record = {"index": index, "skip": True, "name": path.name}
                        else:
                            relative_folder = ""
                            if source_root is not None:
                                try:
                                    relative_folder = path.parent.relative_to(Path(source_root).resolve()).as_posix()
                                except (ValueError, OSError):
                                    relative_folder = ""
                            prepared = self.library.prepare_import(
                                path, relative_folder=relative_folder, known_digests=known_digests
                            )
                            record = {
                                "index": index,
                                "skip": bool(prepared.get("duplicate")),
                                "prepared": prepared,
                                "name": path.name,
                            }
                    except Exception as exc:
                        record = {"index": index, "error": str(exc), "name": original_path.name}
                    publish_result(record)
                publish_result({"done": True, "cancelled": cancel_event.is_set()}, allow_after_cancel=True)
            except Exception as exc:
                publish_result(
                    {"done": True, "fatal_error": str(exc), "cancelled": cancel_event.is_set()},
                    allow_after_cancel=True
                )

        worker = threading.Thread(target=import_worker, name="RecipeLibraryImport", daemon=True)
        worker.start()

        def finish_import(record=None):
            try:
                self._pop_embedded_page(progress_window)
            except tk.TclError:
                pass
            self._import_cancel_event = None
            self.refresh()
            summary = (
                f"הייבוא הסתיים: {state['added']} נוספו, "
                f"{state['skipped']} דולגו, {state['failed']} נכשלו."
            )
            if cancel_event.is_set():
                summary = (
                    f"הייבוא הופסק: {state['added']} נוספו, "
                    f"{state['skipped']} דולגו, {state['failed']} נכשלו."
                )
            self.status_value.set(summary)
            if state["failed"] or (record and record.get("fatal_error")):
                detail = f"{state['failed']} קבצים לא נקלטו."
                if record and record.get("fatal_error"):
                    detail += "\n" + str(record["fatal_error"])
                messagebox.showwarning(
                    "ייבוא הסתיים עם שגיאות",
                    f"נוספו {state['added']} קבצים.\nדולגו {state['skipped']} קבצים.\n{detail}",
                    parent=self.root,
                )

        def poll_results():
            try:
                record = results.get_nowait()
            except queue.Empty:
                if worker.is_alive():
                    progress_window.after(120, poll_results)
                else:
                    finish_import()
                return

            if record.get("done"):
                finish_import(record)
                return

            state["processed"] += 1
            if record.get("error"):
                state["failed"] += 1
            elif record.get("skip"):
                state["skipped"] += 1
            else:
                try:
                    added, _message = self.library.import_prepared(record["prepared"])
                    if added:
                        state["added"] += 1
                    else:
                        state["skipped"] += 1
                except Exception:
                    state["failed"] += 1

            progress.configure(value=state["processed"])
            counter.set(f"{state['processed']} מתוך {len(paths)}")
            progress_text.set(f"מעבד: {record.get('name', '')[:65]}")
            progress_window.after(20, poll_results)

        progress_window.after(100, poll_results)


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
        """Open a file preview; Office extraction and conversion run off the UI thread."""
        path = Path(path)
        window = self._push_embedded_page(title)
        suffix = path.suffix.lower()
        is_office = suffix in OFFICE_TYPES
        office_state = {
            "loading": is_office,
            "preview_path": path,
            "temporary_directory": None,
            "mode": "source",
            "closed": threading.Event(),
        }
        office_events: queue.Queue = queue.Queue()
        document_state = {"document": None}
        content_host = tk.Frame(window, bg=self.colors["background"])
        content_host.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        extracted_text = content or ""
        if not extracted_text and not is_office:
            try:
                extracted_text = extract_text(path)
            except Exception as exc:
                extracted_text = f"לא ניתן לחלץ טקסט מהקובץ: {exc}"

        def clear_content():
            document = document_state.get("document")
            if document is not None:
                try:
                    document.close()
                except Exception:
                    pass
                document_state["document"] = None
            for child in content_host.winfo_children():
                child.destroy()

        def close_preview():
            office_state["closed"].set()
            clear_content()
            temporary_directory = office_state.get("temporary_directory")
            if temporary_directory is not None:
                shutil.rmtree(temporary_directory, ignore_errors=True)
                office_state["temporary_directory"] = None
            self._pop_embedded_page(window)

        self._embedded_header(
            window, title, "תצוגת המקור או הטקסט שחולץ מהקובץ.",
            back_command=close_preview
        )

        controls = tk.Frame(window, bg=self.colors["background"], padx=20, pady=(0, 9))
        controls.pack(fill="x")
        notice_label = ttk.Label(
            controls, foreground="#9A4B18",
            background="#FFF4E5", padding=8, wraplength=1100, justify="right"
        )

        progress_frame = ttk.Frame(controls, style="App.TFrame")
        progress_text = tk.StringVar(value="מכין את מנוע Office בפעם הראשונה…")
        progress_label = ttk.Label(progress_frame, textvariable=progress_text, style="Muted.TLabel", anchor="e")
        progress_label.pack(fill="x", pady=(0, 4))
        progress_bar = ttk.Progressbar(progress_frame, mode="indeterminate")
        progress_bar.pack(fill="x")
        if is_office:
            progress_frame.pack(fill="x", pady=(0, 8))
            progress_bar.start(12)

        def set_notice(value: str):
            if value:
                notice_label.configure(text=value)
                if not notice_label.winfo_manager():
                    notice_label.pack(fill="x", pady=(0, 8))
            else:
                notice_label.pack_forget()

        def update_office_progress(phase: str, current: int = 0, total: int = 0):
            if office_state["closed"].is_set():
                return
            try:
                if phase == "extracting":
                    progress_bar.stop()
                    progress_bar.configure(mode="determinate", maximum=max(1, total), value=current)
                    progress_text.set(f"מחלץ רכיבי Office: {current:,} מתוך {total:,} קבצים")
                elif phase == "converting":
                    progress_bar.stop()
                    progress_bar.configure(mode="indeterminate")
                    progress_bar.start(12)
                    progress_text.set("ממיר את המסמך לתצוגה…")
                else:
                    progress_bar.stop()
                    progress_bar.configure(mode="indeterminate")
                    progress_bar.start(12)
                    progress_text.set("בודק ומכין את מנוע Office…")
            except tk.TclError:
                pass

        def render_preview(mode):
            office_state["mode"] = mode
            clear_content()
            if mode == "text":
                self._show_text_in_viewer(
                    content_host, extracted_text or "לא קיים טקסט שחולץ מהקובץ."
                )
                text_button.set_active(True)
                source_button.set_active(False)
                return

            text_button.set_active(False)
            source_button.set_active(True)
            if is_office and office_state["loading"]:
                ttk.Label(
                    content_host,
                    text="מכין את מנוע Office ואת תצוגת המסמך…\nאפשר להמשיך להשתמש בחלונית הטקסט בזמן ההכנה.",
                    style="Panel.TLabel", justify="center"
                ).pack(expand=True, pady=36)
                return

            preview_path = office_state["preview_path"]
            if suffix in IMAGE_TYPES and Image is not None:
                self._show_image_in_viewer(content_host, path)
            elif (
                suffix == ".pdf" or suffix in EBOOK_TYPES
                or (is_office and preview_path != path)
            ) and fitz is not None and Image is not None and ImageTk is not None:
                document_state["document"] = self._show_pdf_in_viewer(content_host, preview_path, window)
            else:
                fallback = notice_label.cget("text") if notice_label.winfo_manager() else ""
                self._show_text_in_viewer(
                    content_host, extracted_text or fallback or "אין לקובץ זה תצוגה חזותית זמינה."
                )

        def office_worker():
            def report(phase, current=0, total=0):
                if office_state["closed"].is_set():
                    return False
                office_events.put({
                    "kind": "progress", "phase": phase,
                    "current": current, "total": total,
                })
                return True

            temporary_directory = None
            try:
                report("checking")
                converted_path, temporary_directory, notice = convert_office_to_pdf(
                    path, progress_callback=report
                )
                try:
                    text_value = content or extract_text(path)
                except Exception as exc:
                    text_value = f"לא ניתן לחלץ טקסט מהקובץ: {exc}"
                if office_state["closed"].is_set():
                    if temporary_directory is not None:
                        shutil.rmtree(temporary_directory, ignore_errors=True)
                    return
                office_events.put({
                    "kind": "done",
                    "preview_path": converted_path or path,
                    "temporary_directory": temporary_directory,
                    "notice": notice,
                    "text": text_value,
                })
            except Exception as exc:
                if temporary_directory is not None:
                    shutil.rmtree(temporary_directory, ignore_errors=True)
                if not office_state["closed"].is_set():
                    office_events.put({
                        "kind": "done", "preview_path": path,
                        "temporary_directory": None,
                        "notice": f"לא ניתן להכין תצוגת Office: {exc}",
                        "text": content or "",
                    })

        def poll_office_events():
            nonlocal extracted_text
            if office_state["closed"].is_set():
                return
            finished = False
            try:
                while True:
                    event = office_events.get_nowait()
                    if event.get("kind") == "progress":
                        update_office_progress(
                            event.get("phase", "checking"),
                            event.get("current", 0), event.get("total", 0)
                        )
                    elif event.get("kind") == "done":
                        office_state["loading"] = False
                        office_state["preview_path"] = event.get("preview_path") or path
                        office_state["temporary_directory"] = event.get("temporary_directory")
                        extracted_text = event.get("text") or ""
                        set_notice(event.get("notice", ""))
                        progress_bar.stop()
                        progress_bar.configure(mode="determinate", maximum=1, value=1)
                        progress_text.set("מנוע Office מוכן.")
                        finished = True
            except queue.Empty:
                pass
            except tk.TclError:
                return
            if finished:
                render_preview(office_state["mode"])
                return
            if office_state["loading"]:
                try:
                    window.after(100, poll_office_events)
                except tk.TclError:
                    return

        source_button = ModernButton(
            controls, text="▧  תמונה / מסמך מקורי",
            command=lambda: render_preview("source"),
            width=190, height=38, bg=self.colors["surface"],
            fg=self.colors["navy"], hover_bg=self.colors["soft"],
            active_bg=self.colors["sidebar_active"], radius=12,
            font=("Segoe UI", 9, "bold"), canvas_bg=self.colors["background"]
        )
        source_button.pack(side="right", padx=(0, 6))
        text_button = ModernButton(
            controls, text="☷  טקסט שחולץ",
            command=lambda: render_preview("text"),
            width=150, height=38, bg=self.colors["surface"],
            fg=self.colors["navy"], hover_bg=self.colors["soft"],
            active_bg=self.colors["sidebar_active"], radius=12,
            font=("Segoe UI", 9, "bold"), canvas_bg=self.colors["background"]
        )
        text_button.pack(side="right", padx=6)

        initial_mode = "source" if (
            is_office or suffix in IMAGE_TYPES or suffix == ".pdf" or suffix in EBOOK_TYPES
        ) else "text"
        render_preview(initial_mode)
        if is_office:
            threading.Thread(
                target=office_worker, name="LibreOfficePreview", daemon=True
            ).start()
            window.after(100, poll_office_events)

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
            width = max(320, parent.winfo_width() - 56)
            height = max(240, parent.winfo_height() - 56)
            with Image.open(path) as source:
                image = ImageOps.exif_transpose(source).copy() if ImageOps is not None else source.copy()
            image.thumbnail(
                (width, height),
                getattr(getattr(Image, "Resampling", Image), "LANCZOS", 1)
            )
            photo = ImageTk.PhotoImage(image)
            label = ttk.Label(parent, image=photo, style="Panel.TLabel")
            label.image = photo
            label.pack(expand=True, fill="both")
        except Exception as exc:
            ttk.Label(parent, text=f"לא ניתן להציג את התמונה: {exc}").pack(expand=True)


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
        dialog = self._push_embedded_page("הגדרות")
        self._embedded_header(dialog, "הגדרות", "התאמת המראה ואפשרויות העבודה של הספרייה.")
        frame = ttk.Frame(dialog, padding=18, style="App.TFrame")
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="הגדרות הספרייה", style="Section.TLabel").pack(fill="x", pady=(0, 12))

        theme_var = tk.StringVar(value=self.library.get_setting("color_theme", "שמנת"))
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

        option_row("ערכת עיצוב:", theme_var, ["שמנת", "מרווה", "לילה", "כחול", "ירוק", "סגול"])
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
            self._pop_embedded_page(dialog)
            self._apply_styles()
            self._rebuild_shell()
            self.status_value.set("ההגדרות נשמרו.")

        buttons = ttk.Frame(frame, style="App.TFrame")
        buttons.pack(fill="x", pady=(14, 0))
        ttk.Button(buttons, text="פתיחת תיקיית הנתונים", command=self.open_data_folder).pack(side="right")
        ttk.Button(buttons, text="שמירה והחלה", style="Primary.TButton", command=save_settings).pack(side="right", padx=8)
        ttk.Button(buttons, text="חזרה", command=lambda: self._pop_embedded_page(dialog)).pack(side="left")

    def close(self):
        active_import = getattr(self, "_import_cancel_event", None)
        if active_import is not None:
            active_import.set()
            messagebox.showinfo(
                "ייבוא פעיל",
                "הייבוא מתבטל. אפשר לסגור את התוכנה לאחר שהקובץ הנוכחי יסיים עיבוד."
            )
            return
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