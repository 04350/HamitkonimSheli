# -*- coding: utf-8 -*-
"""חלון כלי עזר נפרד לספריית המתכונים: הדפסה, קניות, תכנון ומצב בישול."""
from __future__ import annotations

import json
import os
import atexit
import re
import shutil
import tempfile
import webbrowser
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt, RGBColor
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except Exception:
    Document = None
    WD_ALIGN_PARAGRAPH = Inches = Pt = RGBColor = OxmlElement = qn = None


def recipe_search_text(title: str, content: str, recipe: dict | None = None, tags=()) -> str:
    """טקסט לחיפוש הכולל גם שדות מובנים של מתכון."""
    recipe = recipe if isinstance(recipe, dict) else {}
    values = [
        title, content, recipe.get("category", ""),
        recipe.get("ingredients", ""), recipe.get("instructions", ""),
        recipe.get("notes", ""), recipe.get("source_folder", ""),
        recipe.get("rating", ""), " ".join(str(tag) for tag in tags)
    ]
    return " ".join(str(value or "") for value in values)




def parse_recipe_text(text: str) -> dict:
    """חילוץ שמרני של שם, מצרכים והוראות מתוך טקסט בעברית."""
    lines = [re.sub(r"\s+", " ", line).strip(" \t•▪●") for line in str(text or "").splitlines()]
    lines = [line for line in lines if line]
    ingredient_header = re.compile(r"^(?:מצרכים|רכיבים|החומרים|רשימת מצרכים)\s*[:：]?$", re.I)
    instruction_header = re.compile(r"^(?:אופן ההכנה|הוראות הכנה|שלבי הכנה|הכנה|אופן הכנת)\s*[:：]?$", re.I)
    other_header = re.compile(r"^(?:הערות|לקישוט|לציפוי|למילוי|רוטב|הגשה|טיפים)\s*[:：]?$", re.I)
    title = ""
    for line in lines[:12]:
        if ingredient_header.match(line) or instruction_header.match(line):
            continue
        if len(line) <= 120 and len(line.split()) <= 16 and not re.fullmatch(r"[\d\W_]+", line):
            title = line
            break
    ingredients, instructions = [], []
    section, found_header = "", False
    for line in lines:
        if ingredient_header.match(line):
            section, found_header = "ingredients", True
            continue
        if instruction_header.match(line):
            section, found_header = "instructions", True
            continue
        if other_header.match(line):
            section = ""
            continue
        if line.endswith(":") and len(line) < 45 and not re.search(r"\d", line):
            if section == "ingredients":
                section = ""
            continue
        if section == "ingredients":
            ingredients.append(line)
        elif section == "instructions":
            instructions.append(line)
    if not ingredients:
        units = ("גרם", "קילו", 'ק"ג', 'מ"ל', "ליטר", "כוס", "כפות", "כף", "כפית", "ביצים", "ביצה", "חבילה", "שקית")
        for line in lines:
            if re.match(r"^(?:\d+(?:[.,]\d+)?(?:\s*/\s*\d+)?|[½¼¾⅓⅔⅛⅜⅝⅞])\s*", line) or any(unit in line for unit in units):
                ingredients.append(line)
    return {
        "title": title,
        "ingredients": "\n".join(ingredients[:100]),
        "instructions": "\n".join(instructions[:160]),
        "is_recipe": bool(found_header or ingredients or instructions),
    }


def is_recipe_record(recipe: dict, content: str = "") -> bool:
    """קובע אם פריט הוא מתכון, בלי להתייחס לכל מסמך כאל מתכון."""
    recipe = recipe if isinstance(recipe, dict) else {}
    if recipe.get("is_recipe") is not None and recipe.get("is_recipe") != "":
        return bool(recipe.get("is_recipe"))
    if any(str(recipe.get(key) or "").strip() for key in ("ingredients", "instructions", "category")):
        return True
    return bool(parse_recipe_text(content)["is_recipe"])


def recipe_missing_parts(recipe: dict) -> list[str]:
    recipe = recipe if isinstance(recipe, dict) else {}
    missing = []
    if not str(recipe.get("ingredients") or "").strip():
        missing.append("מצרכים")
    if not str(recipe.get("instructions") or "").strip():
        missing.append("הוראות הכנה")
    return missing

def compile_shopping_list(recipes) -> str:
    """מאחד שורות מצרכים זהות בלבד, בלי לחבר כמויות או יחידות שונות."""
    lines = []
    seen = set()
    for recipe in recipes or []:
        ingredients = str(recipe.get("ingredients") or "").strip()
        for raw in ingredients.splitlines():
            line = raw.strip().lstrip("•▪●*- ").strip()
            if not line or len(line) > 220:
                continue
            key = re.sub(r"\s+", " ", line).casefold()
            if key not in seen:
                seen.add(key)
                lines.append("☐ " + line)
    return "\n".join(lines)


def build_recipe_book_docx(
    recipes, title: str, target: Path, include_notes: bool = True,
    template: str = "קלאסי", layout: str = "עמוד לכל מתכון",
    include_toc: bool = True, include_images: bool = True,
    include_page_numbers: bool = True,
):
    """יוצר חוברת Word בעברית, עם עיצובים, תמונות, תוכן עניינים ומספור עמודים."""
    if Document is None or WD_ALIGN_PARAGRAPH is None:
        raise RuntimeError("רכיב יצירת מסמכי Word אינו זמין.")
    recipes = list(recipes or [])
    if not recipes:
        raise ValueError("לא נבחרו מתכונים להדפסה.")
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.55 if layout != "עמוד לכל מתכון" else 0.65)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.68)
    section.right_margin = Inches(0.68)
    themes = {
        "קלאסי": {"title": RGBColor(31, 54, 78), "accent": RGBColor(31, 54, 78), "scale": 1.0},
        "אלגנטי": {"title": RGBColor(103, 75, 112), "accent": RGBColor(150, 116, 87), "scale": 1.0},
        "פשוט": {"title": RGBColor(25, 25, 25), "accent": RGBColor(90, 90, 90), "scale": 0.96},
    }
    theme = themes.get(template, themes["קלאסי"])
    scale = theme["scale"]
    if layout == "קומפקטי":
        scale *= 0.86
    elif layout == "כרטיסים קטנים (2 בעמוד)":
        scale *= 0.72
    normal_style = document.styles["Normal"]
    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(11 * scale)
    normal_rpr = normal_style._element.get_or_add_rPr()
    normal_fonts = normal_rpr.rFonts
    if normal_fonts is None:
        normal_fonts = OxmlElement("w:rFonts")
        normal_rpr.insert(0, normal_fonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        normal_fonts.set(qn(f"w:{key}"), "Arial")

    def add_line(text, size=11, bold=False, align=None, space_after=5, color=None):
        paragraph = document.add_paragraph()
        paragraph.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.RIGHT
        paragraph.paragraph_format.space_after = Pt(max(1, space_after * scale))
        paragraph.paragraph_format.keep_together = True
        ppr = paragraph._p.get_or_add_pPr()
        bidi = OxmlElement("w:bidi")
        bidi.set(qn("w:val"), "1")
        ppr.append(bidi)
        run = paragraph.add_run(str(text or ""))
        run.font.name = "Arial"
        run.font.size = Pt(max(7, size * scale))
        run.bold = bold
        if color is not None:
            run.font.color.rgb = color
        rpr = run._r.get_or_add_rPr()
        rtl = OxmlElement("w:rtl")
        rtl.set(qn("w:val"), "1")
        rpr.append(rtl)
        fonts = rpr.rFonts
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            rpr.insert(0, fonts)
        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(qn(f"w:{key}"), "Arial")
        return paragraph

    if include_page_numbers:
        footer = section.footer.paragraphs[0]
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = footer.add_run()
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instruction = OxmlElement("w:instrText")
        instruction.set(qn("xml:space"), "preserve")
        instruction.text = " PAGE "
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        number = OxmlElement("w:t")
        number.text = "1"
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        run._r.extend([begin, instruction, separate, number, end])

    add_line(title or "חוברת המתכונים שלי", 26, True, WD_ALIGN_PARAGRAPH.CENTER, 10, theme["title"])
    add_line(f"חוברת מתכונים · {len(recipes)} פריטים", 11, False, WD_ALIGN_PARAGRAPH.CENTER, 8, theme["accent"])
    add_line("נוצרה מתוך ספריית המתכונים שלי", 9, False, WD_ALIGN_PARAGRAPH.CENTER, 0)
    document.add_page_break()
    if include_toc:
        add_line("תוכן החוברת", 20, True, space_after=12, color=theme["title"])
        for index, recipe in enumerate(recipes, 1):
            first_recipe_page = 3 if include_toc else 2
            if layout == "עמוד לכל מתכון":
                page_number = index + first_recipe_page - 1
            elif layout == "כרטיסים קטנים (2 בעמוד)":
                page_number = ((index - 1) // 2) + first_recipe_page
            else:
                page_number = "—"
            add_line(f"{index}. {recipe.get('title') or f'מתכון {index}'}  ·  עמוד {page_number}", 11, False, space_after=4)
        document.add_page_break()

    for index, recipe in enumerate(recipes):
        if index and layout == "עמוד לכל מתכון":
            document.add_page_break()
        elif index and layout == "כרטיסים קטנים (2 בעמוד)" and index % 2 == 0:
            document.add_page_break()
        add_line(recipe.get("title") or f"מתכון {index + 1}", 21, True, space_after=7, color=theme["title"])
        metadata = []
        for key, label, suffix in (
            ("category", "קטגוריה", ""),
            ("prep_minutes", "זמן הכנה", " דקות"),
            ("cook_minutes", "זמן בישול", " דקות"),
            ("servings", "כמות", " מנות"),
            ("rating", "דירוג", " ★"),
        ):
            value = str(recipe.get(key) or "").strip()
            if value and value not in ("0", "ללא דירוג"):
                metadata.append(f"{label}: {value}{suffix}")
        if metadata:
            add_line("  |  ".join(metadata), 9, False, space_after=8, color=theme["accent"])
        image_path = str(recipe.get("image_path") or "").strip()
        if include_images and image_path and Path(image_path).is_file():
            try:
                document.add_picture(image_path, width=Inches(3.3 if layout != "כרטיסים קטנים (2 בעמוד)" else 2.2))
                document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                pass
        ingredients = str(recipe.get("ingredients") or "").strip()
        instructions = str(recipe.get("instructions") or "").strip()
        notes = str(recipe.get("notes") or "").strip()
        source = str(recipe.get("content") or "").strip()
        if ingredients:
            add_line("מצרכים", 14, True, space_after=4, color=theme["accent"])
            for line in ingredients.splitlines():
                if line.strip():
                    add_line("• " + line.strip(), 10 if layout == "כרטיסים קטנים (2 בעמוד)" else 11, False, space_after=2)
        if instructions:
            add_line("הוראות הכנה", 14, True, space_after=4, color=theme["accent"])
            for line in instructions.splitlines():
                if line.strip():
                    add_line(line.strip(), 10 if layout == "כרטיסים קטנים (2 בעמוד)" else 11, False, space_after=3)
        if include_notes and notes:
            add_line("הערות אישיות", 13, True, space_after=4, color=theme["accent"])
            for line in notes.splitlines():
                if line.strip():
                    add_line(line.strip(), 9, False, space_after=2)
        if source:
            known_lines = {
                re.sub(r"\s+", " ", line.strip()).casefold()
                for block in (ingredients, instructions)
                for line in block.splitlines() if line.strip()
            }
            remaining_lines = [
                line.strip() for line in source.splitlines()
                if line.strip() and re.sub(r"\s+", " ", line.strip()).casefold() not in known_lines
            ]
            if remaining_lines:
                add_line("טקסט נוסף מהקובץ המקורי", 12, True, space_after=4, color=theme["accent"])
                for line in remaining_lines:
                    add_line(line, 9, False, space_after=2)
        elif not ingredients and not instructions and not notes:
            add_line("לא הוזנו פרטי מתכון לפריט זה.", 10)
        if layout == "כרטיסים קטנים (2 בעמוד)" and index % 2 == 0 and index + 1 < len(recipes):
            add_line("─" * 48, 8, False, WD_ALIGN_PARAGRAPH.CENTER, 5, theme["accent"])
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(target))
    return target

def export_recipe_book_pdf(recipes, title: str, target: Path, converter, include_notes: bool = True, template: str = 'קלאסי', layout: str = 'עמוד לכל מתכון', include_toc: bool = True, include_images: bool = True, include_page_numbers: bool = True) -> Path:
    """ממיר חוברת Word ל-PDF בעזרת מנוע LibreOffice שכבר כלול בחבילת Windows."""
    target = Path(target)
    if target.suffix.lower() != ".pdf":
        target = target.with_suffix(".pdf")
    temporary_root = Path(tempfile.mkdtemp(prefix="recipe-book-"))
    conversion_directory = None
    try:
        docx_path = temporary_root / "recipe-book.docx"
        build_recipe_book_docx(recipes, title, docx_path, include_notes, template, layout, include_toc, include_images, include_page_numbers)
        pdf_path, conversion_directory, error = converter(docx_path, timeout=240)
        if pdf_path is None:
            raise RuntimeError(error or "לא ניתן ליצור קובץ PDF.")
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(pdf_path, target)
        if not target.is_file() or target.stat().st_size < 100:
            raise RuntimeError("קובץ ה-PDF שנוצר אינו תקין.")
        return target
    finally:
        if conversion_directory:
            shutil.rmtree(conversion_directory, ignore_errors=True)
        shutil.rmtree(temporary_root, ignore_errors=True)


class RecipeCenter:
    """חלון נפרד, כדי שכלי המתכונים לא יעמיסו על מסך ניהול הספרייה הראשי."""

    def __init__(self, app, extract_ingredients, file_kind, safe_title, converter):
        self.app = app
        self.extract_ingredients = extract_ingredients
        self.file_kind = file_kind
        self.safe_title = safe_title
        self.converter = converter
        self.records = []
        self.by_id = {}
        self.preview_dirs = []

    def show(self, initial_item_id=None):
        app = self.app
        window = app._push_embedded_page("מרכז המתכונים")
        app._embedded_header(
            window, "מרכז המתכונים",
            "הדפסה, רשימות קניות, תכנון שבועי, כרטיסי מתכון ומצב בישול."
        )
        outer = ttk.Frame(window, padding=16, style="App.TFrame")
        outer.pack(fill="both", expand=True)
        ttk.Label(
            outer, text="כלי עזר לעבודה עם המתכונים — בלי לשנות את מסך הספרייה וניהול הקבצים.",
            style="Muted.TLabel"
        ).pack(fill="x", anchor="e", pady=(3, 12))

        rows = list(app.library.rows())
        rows.sort(key=lambda row: str(row["title"]).casefold())
        self.records = []
        self.by_id = {}
        for row in rows:
            data = app.library.get_recipe_data(row["id"])
            record = {
                "id": int(row["id"]), "title": str(row["title"]),
                "category": str(data.get("category") or ""),
                "ingredients": str(data.get("ingredients") or "") or self.extract_ingredients(row["content"] or ""),
                "instructions": str(data.get("instructions") or ""),
                "prep_minutes": str(data.get("prep_minutes") or ""),
                "cook_minutes": str(data.get("cook_minutes") or ""),
                "servings": str(data.get("servings") or ""),
                "notes": str(data.get("notes") or ""),
                "content": str(row["content"] or ""),
                "extension": str(row["extension"] or ""),
                "image_path": str(data.get("image_path") or (row["stored_path"] if str(row["extension"] or "").lower() in {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"} else "")),
                "source_folder": str(data.get("source_folder") or ""),
                "rating": str(data.get("rating") or ""),
                "is_recipe": is_recipe_record(data, row["content"] or ""),
                "history": list(data.get("history") or []),
            }
            self.records.append(record)
            self.by_id[record["id"]] = record

        if not self.records:
            ttk.Label(outer, text="הספרייה עדיין ריקה. הוסיפו קבצים או מתכונים כדי להתחיל.", style="Muted.TLabel").pack(fill="both", expand=True)
            ttk.Button(outer, text="חזרה לספרייה", command=lambda: app._pop_embedded_page(window)).pack(anchor="w")
            return window

        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True)

        def make_selector(parent, height=16):
            frame = ttk.Frame(parent, padding=10, style="Card.TFrame")
            ttk.Label(frame, text="בחירת מתכונים ופריטים מהספרייה", style="Section.TLabel").pack(fill="x", anchor="e", pady=(0, 5))
            ttk.Label(frame, text="לבחירת כמה פריטים: Ctrl או Shift.", style="Muted.TLabel").pack(fill="x", anchor="e", pady=(0, 7))
            area = ttk.Frame(frame, style="Card.TFrame")
            area.pack(fill="both", expand=True)
            box = tk.Listbox(
                area, selectmode=tk.EXTENDED, exportselection=False, justify="right",
                activestyle="none", height=height, font=("Segoe UI", 10),
                background="#FFFFFF", foreground=app.colors["text"],
                selectbackground=app.colors["selected"], selectforeground=app.colors["navy"],
                relief="flat", highlightthickness=1, highlightbackground=app.colors["border"]
            )
            scroll = ttk.Scrollbar(area, orient="vertical", command=box.yview)
            box.configure(yscrollcommand=scroll.set)
            box.pack(side="right", fill="both", expand=True)
            scroll.pack(side="left", fill="y")
            for record in self.records:
                box.insert("end", f"{record['title']}  ·  {self.file_kind(record['extension'])}")
            quick = ttk.Frame(frame, style="Card.TFrame")
            quick.pack(fill="x", pady=(7, 0))
            ttk.Button(quick, text="בחירת הכול", command=lambda b=box: b.selection_set(0, "end")).pack(side="right", padx=4)
            ttk.Button(quick, text="ניקוי בחירה", command=lambda b=box: b.selection_clear(0, "end")).pack(side="right")
            return frame, box

        def selected(box):
            return [self.records[index] for index in box.curselection()]

        # 1. PDF printing and booklets.
        print_tab = ttk.Frame(notebook, padding=12, style="App.TFrame")
        notebook.add(print_tab, text="🖨 הדפסה וחוברות")
        selector_frame, print_box = make_selector(print_tab)
        selector_frame.pack(side="right", fill="both", expand=True, padx=(8, 0))
        if initial_item_id is not None:
            for index, record in enumerate(self.records):
                if record["id"] == int(initial_item_id):
                    print_box.selection_set(index)
                    print_box.see(index)
                    break
        options = ttk.Frame(print_tab, padding=16, style="Card.TFrame")
        options.pack(side="left", fill="y", padx=(0, 8))
        ttk.Label(options, text="יצירת חוברת להדפסה", style="Section.TLabel").pack(fill="x", anchor="e", pady=(0, 8))
        ttk.Label(
            options, text="בחרו מתכון אחד או כמה מתכונים. כל מתכון יתחיל בעמוד חדש. הקובץ ייווצר כ-PDF בעברית.",
            style="Muted.TLabel", wraplength=260, justify="right"
        ).pack(fill="x", pady=(0, 12))
        title_var = tk.StringVar(value="חוברת המתכונים שלי")
        ttk.Label(options, text="כותרת החוברת:", style="Panel.TLabel").pack(fill="x", anchor="e")
        ttk.Entry(options, textvariable=title_var, justify="right", width=30).pack(fill="x", pady=(4, 12))
        notes_var = tk.BooleanVar(value=True)
        open_var = tk.BooleanVar(value=True)
        toc_var = tk.BooleanVar(value=True)
        images_var = tk.BooleanVar(value=True)
        pages_var = tk.BooleanVar(value=True)
        template_var = tk.StringVar(value="קלאסי")
        layout_var = tk.StringVar(value="עמוד לכל מתכון")
        ttk.Label(options, text="עיצוב החוברת:", style="Panel.TLabel").pack(fill="x", anchor="e")
        ttk.Combobox(options, textvariable=template_var, values=["קלאסי", "אלגנטי", "פשוט"], state="readonly", width=26, justify="right").pack(fill="x", pady=(3, 7))
        ttk.Label(options, text="פריסת עמודים:", style="Panel.TLabel").pack(fill="x", anchor="e")
        ttk.Combobox(options, textvariable=layout_var, values=["עמוד לכל מתכון", "קומפקטי", "כרטיסים קטנים (2 בעמוד)"], state="readonly", width=26, justify="right").pack(fill="x", pady=(3, 7))
        ttk.Checkbutton(options, text="לכלול הערות אישיות", variable=notes_var).pack(anchor="e", pady=2)
        ttk.Checkbutton(options, text="לכלול תוכן החוברת", variable=toc_var).pack(anchor="e", pady=2)
        ttk.Checkbutton(options, text="לכלול תמונות", variable=images_var).pack(anchor="e", pady=2)
        ttk.Checkbutton(options, text="מספרי עמודים", variable=pages_var).pack(anchor="e", pady=2)
        ttk.Checkbutton(options, text="לפתוח PDF לאחר היצירה", variable=open_var).pack(anchor="e", pady=2)

        def export_options():
            return {
                "include_notes": notes_var.get(),
                "template": template_var.get(),
                "layout": layout_var.get(),
                "include_toc": toc_var.get(),
                "include_images": images_var.get(),
                "include_page_numbers": pages_var.get(),
            }

        def selected_recipes():
            chosen = selected(print_box)
            for record in chosen:
                source_record = self.by_id.get(record["id"], {})
                record["image_path"] = source_record.get("image_path", "")
                record["source_folder"] = source_record.get("source_folder", "")
                record["rating"] = source_record.get("rating", "")
            return chosen

        def preview_pdf():
            chosen = selected_recipes()
            if not chosen:
                messagebox.showinfo("תצוגה מקדימה", "בחרו לפחות מתכון אחד.", parent=window)
                return
            try:
                preview_dir = Path(tempfile.mkdtemp(prefix="recipe-print-preview-"))
                self.preview_dirs.append(preview_dir)
                atexit.register(shutil.rmtree, preview_dir, ignore_errors=True)
                preview_path = preview_dir / "תצוגה-מקדימה.pdf"
                export_recipe_book_pdf(chosen, title_var.get().strip() or "חוברת המתכונים שלי", preview_path, self.converter, **export_options())
                app.show_viewer(preview_path, "תצוגה מקדימה – חוברת מתכונים")
            except Exception as exc:
                messagebox.showerror("תצוגה מקדימה", f"לא ניתן ליצור תצוגה מקדימה.\\n{exc}", parent=window)

        def create_pdf():
            chosen = selected_recipes()
            if not chosen:
                messagebox.showinfo("יצירת חוברת", "בחרו לפחות פריט אחד מהרשימה.", parent=window)
                return
            title = title_var.get().strip() or "חוברת המתכונים שלי"
            name = self.safe_title(title) or "חוברת המתכונים שלי"
            target = filedialog.asksaveasfilename(
                parent=window, title="שמירת חוברת מתכונים כ-PDF",
                defaultextension=".pdf", initialfile=name + ".pdf",
                filetypes=[("קובץ PDF", "*.pdf")]
            )
            if not target:
                return
            try:
                result = export_recipe_book_pdf(chosen, title, Path(target), self.converter, **export_options())
                app.status_value.set(f"חוברת PDF נוצרה: {result.name}")
                messagebox.showinfo("החוברת מוכנה", f"קובץ PDF נוצר בהצלחה.\n\n{result}\n\nאפשר לפתוח אותו ולהדפיס במדפסת הרגילה.", parent=window)
                if open_var.get():
                    app.show_viewer(result, "חוברת המתכונים – תצוגת PDF")
            except Exception as exc:
                messagebox.showerror("שגיאה ביצירת PDF", f"לא ניתן ליצור את החוברת.\n{exc}\n\nבדקו שתיקיית LibreOffice נמצאת לצד קובץ התוכנה.", parent=window)

        ttk.Button(options, text="תצוגה מקדימה לפני יצירה", command=preview_pdf).pack(fill="x", pady=(12, 5))
        ttk.Button(options, text="יצירת חוברת PDF  🖨", style="Primary.TButton", command=create_pdf).pack(fill="x", pady=(3, 6))
        ttk.Label(options, text="ה-PDF הוא קובץ עצמאי שאפשר להדפיס, לשלוח או לשמור.", style="Muted.TLabel", wraplength=260, justify="right").pack(fill="x", pady=(6, 0))

        # 2. Shopping list from multiple recipes.
        shopping_tab = ttk.Frame(notebook, padding=12, style="App.TFrame")
        notebook.add(shopping_tab, text="🛒 רשימת קניות")
        shopping_selector, shopping_box = make_selector(shopping_tab, 12)
        shopping_selector.pack(side="right", fill="both", expand=True, padx=(8, 0))
        result_frame = ttk.Frame(shopping_tab, padding=14, style="Card.TFrame")
        result_frame.pack(side="left", fill="both", expand=True, padx=(0, 8))
        ttk.Label(result_frame, text="רשימת הקניות", style="Section.TLabel").pack(fill="x", anchor="e")
        ttk.Label(
            result_frame, text="רכיבים זהים יאוחדו. כמויות או יחידות שונות יישארו נפרדות כדי למנוע חיבור שגוי.",
            style="Muted.TLabel", wraplength=410, justify="right"
        ).pack(fill="x", pady=(3, 8))
        shopping_text = tk.Text(result_frame, wrap="word", font=("Segoe UI", 11), background="#FFFFFF", foreground=app.colors["text"], relief="flat", padx=12, pady=10)
        shopping_text.pack(fill="both", expand=True)
        shopping_text.tag_configure("rtl", justify="right")
        shopping_text.insert("1.0", "בחרו מתכונים בצד ולחצו על יצירת רשימת קניות.")
        shopping_text.tag_add("rtl", "1.0", "end")

        def generate_shopping():
            chosen = selected(shopping_box)
            if not chosen:
                messagebox.showinfo("רשימת קניות", "בחרו לפחות מתכון אחד.", parent=window)
                return
            content = compile_shopping_list(chosen) or "לא נמצאו מצרכים. אפשר להוסיף אותם בחלון ״תצוגת מתכון מקצועית״."
            shopping_text.delete("1.0", "end")
            shopping_text.insert("1.0", content)
            shopping_text.tag_add("rtl", "1.0", "end")

        def save_shopping():
            content = shopping_text.get("1.0", "end-1c").strip()
            if not content or content.startswith("בחרו מתכונים"):
                messagebox.showinfo("רשימת קניות", "קודם צרו רשימת קניות.", parent=window)
                return
            target = filedialog.asksaveasfilename(parent=window, title="שמירת רשימת קניות", defaultextension=".txt", initialfile="רשימת-קניות.txt", filetypes=[("קובץ טקסט", "*.txt")])
            if target:
                try:
                    Path(target).write_text(content + "\n", encoding="utf-8")
                    messagebox.showinfo("רשימת קניות", f"הרשימה נשמרה.\n{target}", parent=window)
                except Exception as exc:
                    messagebox.showerror("רשימת קניות", str(exc), parent=window)

        buttons = ttk.Frame(result_frame, style="Card.TFrame")
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="יצירת רשימה מהמתכונים המסומנים", style="Primary.TButton", command=generate_shopping).pack(side="right", padx=4)
        ttk.Button(buttons, text="שמירה כקובץ טקסט", command=save_shopping).pack(side="left", padx=4)

        # 3. Weekly planner, persisted in the existing local settings table.
        weekly_tab = ttk.Frame(notebook, padding=16, style="App.TFrame")
        notebook.add(weekly_tab, text="📅 תכנון שבועי")
        ttk.Label(weekly_tab, text="תכנון ארוחות לשבוע", style="Section.TLabel").pack(fill="x", anchor="e")
        ttk.Label(weekly_tab, text="בחירת מתכון לכל ארוחה. התכנון נשמר במחשב גם לאחר סגירת התוכנה.", style="Muted.TLabel").pack(fill="x", anchor="e", pady=(3, 12))
        plan_frame = ttk.Frame(weekly_tab, padding=12, style="Card.TFrame")
        plan_frame.pack(fill="both", expand=True)
        display_by_id = {record["id"]: f"#{record['id']} · {record['title']}" for record in self.records}
        id_by_display = {value: key for key, value in display_by_id.items()}
        menu_values = ["— לא נבחר —"] + list(display_by_id.values())
        headers = [("ארוחת ערב", 0), ("ארוחת צהריים", 1), ("ארוחת בוקר", 2), ("יום", 3)]
        for label, column in headers:
            ttk.Label(plan_frame, text=label, style="Panel.TLabel").grid(row=0, column=column, padx=5, pady=(0, 8), sticky="ew")
        for column in range(4):
            plan_frame.grid_columnconfigure(column, weight=1)
        days = ["יום ראשון", "יום שני", "יום שלישי", "יום רביעי", "יום חמישי", "יום שישי", "שבת"]
        meals = [("בוקר", 2), ("צהריים", 1), ("ערב", 0)]
        variables = {}
        try:
            saved = json.loads(app.library.get_setting("weekly_menu", "{}") or "{}")
            if not isinstance(saved, dict):
                saved = {}
        except (ValueError, TypeError):
            saved = {}
        for row_index, day in enumerate(days, start=1):
            ttk.Label(plan_frame, text=day, style="Panel.TLabel").grid(row=row_index, column=3, padx=5, pady=5, sticky="e")
            variables[day] = {}
            saved_day = saved.get(day, {})
            for meal, column in meals:
                variable = tk.StringVar(value="— לא נבחר —")
                try:
                    saved_id = int(saved_day.get(meal)) if isinstance(saved_day, dict) and saved_day.get(meal) is not None else None
                except (ValueError, TypeError):
                    saved_id = None
                if saved_id is not None:
                    variable.set(display_by_id.get(saved_id, "— לא נבחר —"))
                ttk.Combobox(plan_frame, textvariable=variable, values=menu_values, state="readonly", width=29, justify="right").grid(row=row_index, column=column, padx=5, pady=5, sticky="ew")
                variables[day][meal] = variable

        def save_plan():
            plan = {}
            for day, day_meals in variables.items():
                plan[day] = {meal: id_by_display[var.get()] for meal, var in day_meals.items() if var.get() in id_by_display}
            app.library.set_setting("weekly_menu", json.dumps(plan, ensure_ascii=False))
            app.status_value.set("תכנון התפריט השבועי נשמר.")
            messagebox.showinfo("תכנון שבועי", "התפריט השבועי נשמר במחשב.", parent=window)

        ttk.Button(weekly_tab, text="שמירת התכנון השבועי", style="Primary.TButton", command=save_plan).pack(anchor="e", pady=(12, 0))

        # 4. כרטיס מתכון עשיר והיסטוריית הכנה.
        card_tab = ttk.Frame(notebook, padding=16, style="App.TFrame")
        notebook.add(card_tab, text="🗂 כרטיס מתכון")
        ttk.Label(card_tab, text="כרטיס מתכון והיסטוריית הכנה", style="Section.TLabel").pack(fill="x", anchor="e")
        card_selector = tk.StringVar(value="— בחירת מתכון —")
        card_combo = ttk.Combobox(card_tab, textvariable=card_selector, values=["— בחירת מתכון —"] + list(display_by_id.values()), state="readonly", width=55, justify="right")
        card_combo.pack(fill="x", pady=(8, 10))
        card_frame = ttk.Frame(card_tab, padding=14, style="Card.TFrame")
        card_frame.pack(fill="both", expand=True)
        card_title = ttk.Label(card_frame, text="בחרו מתכון להצגת הכרטיס.", style="Section.TLabel")
        card_title.pack(fill="x", anchor="e")
        card_image_label = ttk.Label(card_frame, style="Panel.TLabel")
        card_image_label.pack(anchor="center", pady=4)
        card_details = tk.Text(card_frame, wrap="word", font=("Segoe UI", 10), height=16, padx=12, pady=10, relief="flat", background="#FFFFFF")
        card_details.pack(fill="both", expand=True, pady=(5, 8))
        card_details.tag_configure("rtl", justify="right")
        self.card_photo = None
        history_frame = ttk.Frame(card_tab, padding=10, style="Card.TFrame")
        history_frame.pack(fill="x")
        ttk.Label(history_frame, text="תיעוד הכנה", style="Panel.TLabel").pack(anchor="e")
        loved_var = tk.BooleanVar(value=True)
        change_var = tk.StringVar()
        ttk.Checkbutton(history_frame, text="אהבתי את התוצאה", variable=loved_var).pack(anchor="e", pady=3)
        ttk.Entry(history_frame, textvariable=change_var, justify="right").pack(fill="x", pady=3)
        change_var.set("מה שיניתי בהכנה הפעם?")
        history_list = tk.Listbox(history_frame, height=4, justify="right", exportselection=False)
        history_list.pack(fill="x", pady=4)

        def refresh_recipe_card(_event=None):
            item_id = id_by_display.get(card_selector.get())
            card_details.configure(state="normal")
            card_details.delete("1.0", "end")
            history_list.delete(0, "end")
            card_image_label.configure(image="", text="")
            self.card_photo = None
            if item_id is None:
                card_title.configure(text="בחרו מתכון להצגת הכרטיס.")
                card_details.configure(state="disabled")
                return
            record = self.by_id[item_id]
            rating = int(record.get("rating") or 0) if str(record.get("rating") or "").isdigit() else 0
            card_title.configure(text=f"{record['title']}  ·  {'★' * rating if rating else 'ללא דירוג'}")
            metadata = [
                f"זמן הכנה: {record.get('prep_minutes') or 'לא צוין'} דקות",
                f"זמן בישול: {record.get('cook_minutes') or 'לא צוין'} דקות",
                f"כמות: {record.get('servings') or 'לא צוין'} מנות",
                f"תיקייה מקורית: {record.get('source_folder') or 'שורש הספרייה'}",
                "",
                "מצרכים:",
                record.get("ingredients") or "לא הוזנו מצרכים",
                "",
                "הוראות הכנה:",
                record.get("instructions") or "לא הוזנו הוראות",
                "",
                "הערות:",
                record.get("notes") or "אין הערות",
            ]
            card_details.insert("1.0", "\n".join(metadata))
            card_details.tag_add("rtl", "1.0", "end")
            card_details.configure(state="disabled")
            image_path = record.get("image_path")
            if image_path:
                try:
                    from PIL import Image, ImageTk
                    with Image.open(image_path) as image:
                        image.thumbnail((220, 150))
                        self.card_photo = ImageTk.PhotoImage(image.copy())
                    card_image_label.configure(image=self.card_photo)
                except Exception:
                    card_image_label.configure(text="לא ניתן להציג את התמונה המצורפת.")
            recipe_data = app.library.get_recipe_data(item_id)
            for entry in (recipe_data.get("history") or [])[-8:]:
                liked = "אהבתי" if entry.get("loved") else "לא אהבתי"
                history_list.insert("end", f"{entry.get('date','')} · {liked} · {entry.get('changes','')}")
        card_combo.bind("<<ComboboxSelected>>", refresh_recipe_card)

        def log_cooking():
            item_id = id_by_display.get(card_selector.get())
            if item_id is None:
                messagebox.showinfo("היסטוריית הכנה", "בחרו מתכון.", parent=window)
                return
            data = app.library.get_recipe_data(item_id)
            history = list(data.get("history") or [])
            changes = change_var.get().strip()
            history.append({
                "date": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "loved": bool(loved_var.get()),
                "changes": changes if changes != "מה שיניתי בהכנה הפעם?" else "",
            })
            data["history"] = history[-100:]
            app.library.set_recipe_data(item_id, data)
            change_var.set("")
            refresh_recipe_card()
            messagebox.showinfo("היסטוריית הכנה", "ההכנה נשמרה בהיסטוריה.", parent=window)
        ttk.Button(history_frame, text="שמירת תיעוד הכנה", style="Primary.TButton", command=log_cooking).pack(anchor="e", pady=(4, 0))

        # 5. מצב בישול מתקדם עם סימון שלבים וטיימרים מקבילים.
        cooking_tab = ttk.Frame(notebook, padding=16, style="App.TFrame")
        notebook.add(cooking_tab, text="🍳 מצב בישול")
        ttk.Label(cooking_tab, text="מצב הכנה — שלב אחר שלב", style="Section.TLabel").pack(fill="x", anchor="e")
        ttk.Label(cooking_tab, text="אפשר לסמן שלבים שהושלמו ולהפעיל כמה טיימרים במקביל.", style="Muted.TLabel").pack(fill="x", anchor="e", pady=(3, 12))
        selector = ttk.Frame(cooking_tab, padding=12, style="Card.TFrame")
        selector.pack(fill="x")
        ttk.Label(selector, text="המתכון:", style="Panel.TLabel").pack(side="right", padx=(0, 8))
        cooking_var = tk.StringVar(value="— בחירת מתכון —")
        cooking_combo = ttk.Combobox(selector, textvariable=cooking_var, values=["— בחירת מתכון —"] + list(display_by_id.values()), state="readonly", width=55, justify="right")
        cooking_combo.pack(side="right", fill="x", expand=True)
        card = ttk.Frame(cooking_tab, padding=22, style="Card.TFrame")
        card.pack(fill="both", expand=True, pady=12)
        count_label = ttk.Label(card, text="בחרו מתכון כדי להתחיל.", style="Muted.TLabel")
        count_label.pack(fill="x", anchor="e", pady=(0, 15))
        step_done = tk.BooleanVar(value=False)
        done_check = ttk.Checkbutton(card, text="סיימתי את השלב הזה", variable=step_done)
        done_check.pack(anchor="e", pady=(0, 8))
        step_label = ttk.Label(card, text="", style="Section.TLabel", justify="right", anchor="ne", wraplength=900)
        step_label.pack(fill="both", expand=True, anchor="ne")
        steps, done_steps, current = [], set(), [0]
        font_size = [22]
        timers = {}

        def render_step():
            if not steps:
                count_label.configure(text="אין שלבי הכנה זמינים.")
                step_label.configure(text="אפשר להוסיף הוראות בחלון ״תצוגת מתכון מקצועית״.")
                done_check.configure(state="disabled")
                return
            done_check.configure(state="normal")
            count_label.configure(text=f"שלב {current[0] + 1} מתוך {len(steps)}")
            step_done.set(current[0] in done_steps)
            step_label.configure(text=steps[current[0]], font=("Segoe UI", font_size[0], "bold"))

        def toggle_done():
            if step_done.get():
                done_steps.add(current[0])
            else:
                done_steps.discard(current[0])
            render_step()
        done_check.configure(command=toggle_done)

        def load_cooking(_event=None):
            item_id = id_by_display.get(cooking_var.get())
            steps.clear()
            done_steps.clear()
            current[0] = 0
            if item_id is not None:
                recipe = self.by_id[item_id]
                raw = recipe.get("instructions") or recipe.get("content") or ""
                steps.extend([line.strip() for line in raw.splitlines() if line.strip()])
            render_step()

        cooking_combo.bind("<<ComboboxSelected>>", load_cooking)
        navigation = ttk.Frame(cooking_tab, style="App.TFrame")
        navigation.pack(fill="x")
        ttk.Button(navigation, text="A−", command=lambda: (font_size.__setitem__(0, max(14, font_size[0]-2)), render_step())).pack(side="left", padx=3)
        ttk.Button(navigation, text="A+", command=lambda: (font_size.__setitem__(0, min(42, font_size[0]+2)), render_step())).pack(side="left", padx=3)
        ttk.Button(navigation, text="השלב הקודם", command=lambda: (current.__setitem__(0, max(0, current[0] - 1)), render_step())).pack(side="right")
        ttk.Button(navigation, text="השלב הבא", style="Primary.TButton", command=lambda: (current.__setitem__(0, min(max(0, len(steps) - 1), current[0] + 1)), render_step())).pack(side="right", padx=8)
        timer_frame = ttk.Frame(cooking_tab, padding=10, style="Card.TFrame")
        timer_frame.pack(fill="x")
        ttk.Label(timer_frame, text="טיימרים פעילים", style="Panel.TLabel").pack(anchor="e")
        timer_rows = ttk.Frame(timer_frame, style="Card.TFrame")
        timer_rows.pack(fill="x", pady=(4, 0))

        def tick_timers():
            now = time.monotonic()
            finished = []
            for timer_id, item in list(timers.items()):
                remaining = max(0, int(item["end"] - now + 0.999))
                minutes, seconds = divmod(remaining, 60)
                item["widget"].configure(text=f"{item['label']}  ·  {minutes:02d}:{seconds:02d}")
                if remaining <= 0:
                    item["widget"].configure(text=f"⏰ הסתיים: {item['label']}")
                    finished.append(timer_id)
                    try:
                        window.bell()
                    except tk.TclError:
                        pass
            for timer_id in finished:
                timers.pop(timer_id, None)
            if timers and window.winfo_exists():
                window.after(1000, tick_timers)

        def start_timer():
            if not steps:
                messagebox.showinfo("טיימר", "בחרו מתכון עם שלבי הכנה.", parent=window)
                return
            match = re.search(r"(\d{1,3})\s*(?:דקות|דקה|min)", steps[current[0]], re.I)
            default_minutes = int(match.group(1)) if match else 5
            minutes = simpledialog.askinteger("טיימר לשלב", "כמה דקות להפעיל עבור השלב הנוכחי?", parent=window, initialvalue=default_minutes, minvalue=1, maxvalue=600)
            if not minutes:
                return
            timer_id = str(time.time_ns())
            timers[timer_id] = {"end": time.monotonic() + minutes * 60, "label": steps[current[0]][:70], "widget": None}
            row = ttk.Frame(timer_rows, style="Card.TFrame")
            row.pack(fill="x", pady=2)
            label = ttk.Label(row, text="", style="Panel.TLabel")
            label.pack(side="right", fill="x", expand=True)
            ttk.Button(row, text="ביטול", command=lambda tid=timer_id, frame=row: (timers.pop(tid, None), frame.destroy())).pack(side="left")
            timers[timer_id]["widget"] = label
            if len(timers) == 1:
                tick_timers()

        ttk.Button(navigation, text="⏱ טיימר לשלב הנוכחי", command=start_timer).pack(side="left", padx=8)
        if initial_item_id is not None and int(initial_item_id) in display_by_id:
            cooking_var.set(display_by_id[int(initial_item_id)])
            load_cooking()
            card_selector.set(display_by_id[int(initial_item_id)])
            refresh_recipe_card()
        return window
