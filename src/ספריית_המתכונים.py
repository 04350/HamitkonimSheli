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
APP_VERSION = "1.0"
IMAGE_TYPES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"}
TEXT_TYPES = {".txt", ".md", ".csv", ".tsv", ".log"}
DOCUMENT_TYPES = {".pdf", ".docx", ".odt", ".rtf", ".html", ".htm", ".xlsx", ".xlsm", ".pptx", ".epub"}
MAX_FILE_SIZE = 250 * 1024 * 1024
MAX_TEXT_LENGTH = 400_000

def configure_ocr():
    """Locate the bundled offline Tesseract engine and Hebrew language data."""
    if pytesseract is None:
        return False
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    candidates = [base / "tesseract" / "tesseract.exe", base / "tesseract.exe"]
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
        if suffix == ".odt":
            names = ["content.xml"]
        elif suffix in (".xlsx", ".xlsm"):
            names = [n for n in archive.namelist() if n == "xl/sharedStrings.xml" or re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)]
        elif suffix == ".pptx":
            names = sorted(n for n in archive.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
        elif suffix == ".epub":
            names = [n for n in archive.namelist() if n.lower().endswith((".xhtml", ".html", ".htm"))]
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
        elif suffix == ".docx" and Document is not None:
            doc = Document(path)
            parts = [p.text for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    parts.append(" | ".join(cell.text for cell in row.cells))
            text = "\n".join(parts)
        elif suffix == ".rtf" and rtf_to_text is not None:
            text = rtf_to_text(decode_text(path.read_bytes()))
        elif suffix in (".odt", ".xlsx", ".xlsm", ".pptx", ".epub"):
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
    if suffix in (".xlsx", ".xlsm"):
        return "גיליון"
    if suffix == ".pptx":
        return "מצגת"
    if suffix == ".epub":
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
        self.db.commit()

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

    def rows(self):
        return self.db.execute("SELECT * FROM items ORDER BY id DESC").fetchall()

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
        self.root.geometry("1180x760")
        self.root.minsize(900, 580)
        self.style = ttk.Style(self.root)
        self._apply_styles()
        self.search_value = tk.StringVar()
        self.status_value = tk.StringVar(value="ברוכים הבאים! אפשר להוסיף קבצים או לגרור אותם לחלון.")
        self.preview_photo = None
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
        """A cohesive, polished Hebrew desktop theme for the whole application."""
        self.colors = {
            "background": "#F3F6FA",
            "surface": "#FFFFFF",
            "navy": "#14263D",
            "text": "#26384D",
            "muted": "#708196",
            "border": "#DCE4EE",
            "accent": "#0F8B8D",
            "accent_hover": "#0B7476",
            "soft": "#EAF3F4",
            "row_alt": "#F7FAFC",
            "selected": "#DDECEF",
            "danger": "#B54747",
        }
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.root.configure(background=self.colors["background"])
        self.style.configure(".", font=("Segoe UI", 10), foreground=self.colors["text"])
        self.style.configure("App.TFrame", background=self.colors["background"])
        self.style.configure("Card.TFrame", background=self.colors["surface"])
        self.style.configure("Panel.TFrame", background=self.colors["surface"])
        self.style.configure("TLabel", background=self.colors["background"], foreground=self.colors["text"])
        self.style.configure("Panel.TLabel", background=self.colors["surface"], foreground=self.colors["text"])
        self.style.configure("Header.TLabel", background=self.colors["background"], foreground=self.colors["navy"], font=("Segoe UI", 25, "bold"))
        self.style.configure("Subheader.TLabel", background=self.colors["background"], foreground=self.colors["muted"], font=("Segoe UI", 10))
        self.style.configure("Status.TLabel", background=self.colors["surface"], foreground=self.colors["muted"], padding=(12, 9), font=("Segoe UI", 9))
        self.style.configure(
            "Card.TLabelframe", background=self.colors["surface"],
            bordercolor=self.colors["border"], borderwidth=1, relief="solid", padding=12
        )
        self.style.configure(
            "Card.TLabelframe.Label", background=self.colors["surface"],
            foreground=self.colors["navy"], font=("Segoe UI", 11, "bold")
        )
        self.style.configure(
            "TButton", font=("Segoe UI", 9, "bold"), padding=(13, 8),
            background="#E8EEF5", foreground=self.colors["navy"],
            bordercolor=self.colors["border"], focuscolor=self.colors["accent"]
        )
        self.style.map(
            "TButton",
            background=[("disabled", "#EEF2F6"), ("pressed", "#D6E0EA"), ("active", "#DDE7F0")],
            foreground=[("disabled", "#9AA7B5"), ("active", self.colors["navy"])]
        )
        self.style.configure(
            "Primary.TButton", background=self.colors["accent"], foreground="#FFFFFF",
            bordercolor=self.colors["accent"], padding=(15, 9)
        )
        self.style.map(
            "Primary.TButton",
            background=[("disabled", "#A8C7C7"), ("pressed", self.colors["accent_hover"]), ("active", self.colors["accent_hover"])],
            foreground=[("disabled", "#F5FAFA"), ("!disabled", "#FFFFFF")]
        )
        self.style.configure(
            "Danger.TButton", background="#FCEEEE", foreground=self.colors["danger"],
            bordercolor="#F2D4D4", padding=(13, 8)
        )
        self.style.map("Danger.TButton", background=[("pressed", "#F5DADA"), ("active", "#F8E3E3")])
        self.style.configure(
            "TEntry", fieldbackground=self.colors["surface"], foreground=self.colors["navy"],
            padding=10, bordercolor=self.colors["border"], insertcolor=self.colors["navy"]
        )
        self.style.configure(
            "Treeview", rowheight=36, font=("Segoe UI", 10),
            background=self.colors["surface"], fieldbackground=self.colors["surface"],
            foreground=self.colors["text"], bordercolor=self.colors["border"], borderwidth=0
        )
        self.style.configure(
            "Treeview.Heading", font=("Segoe UI", 9, "bold"),
            background="#EAF0F6", foreground=self.colors["navy"],
            padding=(10, 11), relief="flat"
        )
        self.style.map(
            "Treeview", background=[("selected", self.colors["selected"])],
            foreground=[("selected", self.colors["navy"])]
        )
        self.style.map("Treeview.Heading", background=[("active", "#DDE7F0")])
        self.style.configure("Vertical.TScrollbar", background="#E2EAF2", troughcolor="#F3F6FA", bordercolor="#F3F6FA", arrowsize=14)

    def _build(self):
        outer = ttk.Frame(self.root, padding=20, style="App.TFrame")
        outer.pack(fill="both", expand=True)
        header = ttk.Frame(outer, style="App.TFrame")
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text=APP_NAME, style="Header.TLabel", anchor="e").pack(fill="x")
        ttk.Label(header, text="כל המתכונים, המסמכים והתמונות שלך — במקום אחד", style="Subheader.TLabel", anchor="e").pack(fill="x", pady=(2, 0))
        tk.Frame(header, height=3, background=self.colors["accent"], bd=0).pack(fill="x", pady=(12, 0))

        search_frame = ttk.Frame(outer, padding=(12, 10), style="Panel.TFrame")
        search_frame.pack(fill="x", pady=(4, 12))
        ttk.Label(search_frame, text="חיפוש", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(side="right", padx=(8, 4))
        entry = ttk.Entry(search_frame, textvariable=self.search_value, justify="right", font=("Segoe UI", 12))
        entry.pack(side="right", fill="x", expand=True, padx=4)
        entry.bind("<KeyRelease>", lambda _event: self.refresh())
        ttk.Button(search_frame, text="נקה חיפוש", command=lambda: self.search_value.set("") or self.refresh()).pack(side="left", padx=5)

        toolbar = ttk.Frame(outer, style="App.TFrame")
        toolbar.pack(fill="x", pady=(0, 10))
        actions = [
            ("הוספת קבצים", self.add_files),
            ("ייבוא תיקייה", self.add_folder),
            ("פתיחת תיקיית הנתונים", self.open_data_folder),
            ("יצירת גיבוי", self.backup),
            ("הגדרות", self.settings),
        ]
        for label, command in actions:
            style = "Primary.TButton" if label == "הוספת קבצים" else "TButton"
            ttk.Button(toolbar, text=label, command=command, style=style).pack(side="right", padx=4)

        body = ttk.Frame(outer, style="App.TFrame")
        body.pack(fill="both", expand=True)
        preview = ttk.LabelFrame(body, text="תצוגה מקדימה", style="Card.TLabelframe")
        preview.pack(side="left", fill="y", padx=(0, 12))
        preview.configure(width=310)
        preview.pack_propagate(False)
        self.preview_image = ttk.Label(preview, text="בחרו תמונה או קובץ מהרשימה", style="Panel.TLabel", anchor="center", justify="center", width=34)
        self.preview_image.pack(fill="both", expand=True)
        self.detail_label = ttk.Label(preview, text="", style="Panel.TLabel", justify="right", anchor="e", wraplength=270)
        self.detail_label.pack(fill="x", pady=10)
        ttk.Button(preview, text="פתיחת הקובץ", command=self.open_selected, style="Primary.TButton").pack(fill="x", pady=4)
        ttk.Button(preview, text="שינוי שם", command=self.rename_selected).pack(fill="x", pady=4)
        ttk.Button(preview, text="הסרה מהספרייה", command=self.delete_selected, style="Danger.TButton").pack(fill="x", pady=4)

        table_frame = ttk.Frame(body, style="Card.TFrame")
        table_frame.pack(side="right", fill="both", expand=True)
        columns = ("date", "kind", "title")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("date", text="תאריך הוספה", anchor="e")
        self.tree.heading("kind", text="סוג קובץ", anchor="e")
        self.tree.heading("title", text="שם המתכון או הקובץ", anchor="e")
        self.tree.column("date", width=135, anchor="e")
        self.tree.column("kind", width=105, anchor="e")
        self.tree.column("title", width=440, anchor="e")
        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.tag_configure("even", background=self.colors["surface"])
        self.tree.tag_configure("odd", background=self.colors["row_alt"])
        self.tree.pack(side="right", fill="both", expand=True, padx=4, pady=4)
        scroll.pack(side="left", fill="y", padx=(0, 3), pady=4)
        self.tree.bind("<<TreeviewSelect>>", self.show_selected)
        self.tree.bind("<Double-1>", lambda _event: self.open_selected())

        ttk.Separator(outer).pack(fill="x", pady=(10, 5))
        ttk.Label(outer, textvariable=self.status_value, style="Status.TLabel", anchor="e").pack(fill="x")

    def refresh(self):
        query = self.search_value.get()
        for item in self.tree.get_children():
            self.tree.delete(item)
        rows = self.library.rows()
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
            self.tree.insert("", "end", iid=iid, values=(row["imported_at"], ext, row["title"]), tags=("even" if index % 2 == 0 else "odd",))
        self.status_value.set(f"נמצאו {len(ranked)} קבצים בספרייה.")
        self.preview_photo = None
        self.preview_image.configure(image="", text="בחרו קובץ מהרשימה")
        self.detail_label.configure(text="")

    def selected_row(self):
        selected = self.tree.selection()
        return self.current_rows.get(selected[0]) if selected else None

    def show_selected(self, _event=None):
        row = self.selected_row()
        if not row:
            return
        path = Path(row["stored_path"])
        info = f"שם: {row['title']}\nקובץ מקורי: {row['original_name']}\nנוסף בתאריך: {row['imported_at']}"
        self.detail_label.configure(text=info)
        self.preview_photo = None
        if path.suffix.lower() in IMAGE_TYPES and Image is not None:
            try:
                with Image.open(path) as image:
                    image.thumbnail((300, 340))
                    self.preview_photo = ImageTk.PhotoImage(image.copy())
                self.preview_image.configure(image=self.preview_photo, text="")
                return
            except Exception:
                pass
        self.preview_image.configure(image="", text="אין תצוגה מקדימה לקובץ הזה")

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
        """Open supported files inside the application, never in an external viewer."""
        window = tk.Toplevel(self.root)
        window.title(title)
        window.geometry("1080x800")
        window.minsize(720, 500)
        window.configure(background=self.colors["background"])
        window.transient(self.root)
        header = ttk.Frame(window, padding=(16, 12), style="Panel.TFrame")
        header.pack(fill="x", padx=12, pady=(12, 4))
        ttk.Label(header, text=title, style="Header.TLabel", anchor="e").pack(fill="x")
        body = ttk.Frame(window, padding=10, style="App.TFrame")
        body.pack(fill="both", expand=True)
        suffix = path.suffix.lower()

        if suffix in IMAGE_TYPES and Image is not None:
            self._show_image_in_viewer(body, path)
        elif suffix == ".pdf" and fitz is not None and Image is not None:
            try:
                doc = fitz.open(path)
                canvas = tk.Canvas(body, background="#E8EDF3", highlightthickness=0)
                scrollbar = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
                inner = ttk.Frame(canvas)
                inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
                canvas.create_window((0, 0), window=inner, anchor="nw")
                canvas.configure(yscrollcommand=scrollbar.set)
                canvas.pack(side="right", fill="both", expand=True)
                scrollbar.pack(side="left", fill="y")
                for number, page in enumerate(doc, 1):
                    pix = page.get_pixmap(matrix=fitz.Matrix(1.35, 1.35), alpha=False)
                    import io
                    image = Image.open(io.BytesIO(pix.tobytes("png")))
                    image.thumbnail((900, 1200))
                    photo = ImageTk.PhotoImage(image.copy())
                    label = ttk.Label(inner, image=photo)
                    label.image = photo
                    label.pack(pady=8)
                    ttk.Label(inner, text=f"עמוד {number}").pack()
            except Exception as exc:
                ttk.Label(body, text=f"לא ניתן להציג את קובץ ה-PDF: {exc}").pack()
        else:
            text = content or extract_text(path)
            if not text:
                text = "אין אפשרות להציג תצוגה פנימית לתוכן הקובץ הזה. אפשר לשמור אותו בספרייה, אך פורמט זה אינו מפוענח כרגע."
            text_frame = ttk.Frame(body)
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

    def _show_image_in_viewer(self, parent, path: Path):
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
