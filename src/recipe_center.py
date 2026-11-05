# -*- coding: utf-8 -*-
"""חלון כלי עזר נפרד לספריית המתכונים: הדפסה, קניות, תכנון ומצב בישול."""
from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except Exception:
    Document = None
    WD_ALIGN_PARAGRAPH = Inches = Pt = OxmlElement = qn = None


def recipe_search_text(title: str, content: str, recipe: dict | None = None, tags=()) -> str:
    """טקסט לחיפוש הכולל גם שדות מובנים של מתכון."""
    recipe = recipe if isinstance(recipe, dict) else {}
    values = [
        title, content, recipe.get("category", ""),
        recipe.get("ingredients", ""), recipe.get("instructions", ""),
        recipe.get("notes", ""), " ".join(str(tag) for tag in tags)
    ]
    return " ".join(str(value or "") for value in values)


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


def build_recipe_book_docx(recipes, title: str, target: Path, include_notes: bool = True):
    """יוצר מסמך Word בעברית ובכיווניות ימין-לשמאל, עמוד נפרד לכל מתכון."""
    if Document is None or WD_ALIGN_PARAGRAPH is None:
        raise RuntimeError("רכיב יצירת מסמכי Word אינו זמין.")
    recipes = list(recipes or [])
    if not recipes:
        raise ValueError("לא נבחרו מתכונים להדפסה.")

    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    normal_style = document.styles["Normal"]
    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(11)
    normal_rpr = normal_style._element.get_or_add_rPr()
    normal_fonts = normal_rpr.rFonts
    if normal_fonts is None:
        normal_fonts = OxmlElement("w:rFonts")
        normal_rpr.insert(0, normal_fonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        normal_fonts.set(qn(f"w:{key}"), "Arial")

    def add_line(text, size=11, bold=False, align=None, space_after=5):
        paragraph = document.add_paragraph()
        paragraph.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.RIGHT
        paragraph.paragraph_format.space_after = Pt(space_after)
        paragraph.paragraph_format.keep_together = True
        ppr = paragraph._p.get_or_add_pPr()
        bidi = OxmlElement("w:bidi")
        bidi.set(qn("w:val"), "1")
        ppr.append(bidi)
        run = paragraph.add_run(str(text or ""))
        run.font.name = "Arial"
        run.font.size = Pt(size)
        run.bold = bold
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

    add_line(title or "חוברת המתכונים שלי", 25, True, WD_ALIGN_PARAGRAPH.CENTER, 10)
    add_line(f"חוברת מתכונים · {len(recipes)} פריטים", 11, False, WD_ALIGN_PARAGRAPH.CENTER, 8)
    add_line("נוצרה מתוך ספריית המתכונים שלי", 9, False, WD_ALIGN_PARAGRAPH.CENTER, 0)
    document.add_page_break()

    for index, recipe in enumerate(recipes):
        if index:
            document.add_page_break()
        add_line(recipe.get("title") or f"מתכון {index + 1}", 21, True, space_after=8)
        metadata = []
        for key, label, suffix in (
            ("category", "קטגוריה", ""),
            ("prep_minutes", "זמן הכנה", " דקות"),
            ("cook_minutes", "זמן בישול", " דקות"),
            ("servings", "כמות", " מנות"),
        ):
            value = str(recipe.get(key) or "").strip()
            if value:
                metadata.append(f"{label}: {value}{suffix}")
        if metadata:
            add_line("  |  ".join(metadata), 9, False, space_after=12)

        ingredients = str(recipe.get("ingredients") or "").strip()
        instructions = str(recipe.get("instructions") or "").strip()
        notes = str(recipe.get("notes") or "").strip()
        source = str(recipe.get("content") or "").strip()

        if ingredients:
            add_line("מצרכים", 14, True, space_after=5)
            for line in ingredients.splitlines():
                if line.strip():
                    add_line("• " + line.strip(), 11, False, space_after=2)
        if instructions:
            add_line("הוראות הכנה", 14, True, space_after=5)
            for line in instructions.splitlines():
                if line.strip():
                    add_line(line.strip(), 11, False, space_after=4)
        if include_notes and notes:
            add_line("הערות אישיות", 14, True, space_after=5)
            for line in notes.splitlines():
                if line.strip():
                    add_line(line.strip(), 10, False, space_after=3)
        if not ingredients and not instructions and source:
            add_line("תוכן המתכון", 14, True, space_after=5)
            for line in source.splitlines():
                if line.strip():
                    add_line(line.strip(), 10, False, space_after=3)
        elif not ingredients and not instructions and not notes and not source:
            add_line("לא הוזנו פרטי מתכון לפריט זה.", 10)

    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(target))
    return target


def export_recipe_book_pdf(recipes, title: str, target: Path, converter, include_notes: bool = True) -> Path:
    """ממיר חוברת Word ל-PDF בעזרת מנוע LibreOffice שכבר כלול בחבילת Windows."""
    target = Path(target)
    if target.suffix.lower() != ".pdf":
        target = target.with_suffix(".pdf")
    temporary_root = Path(tempfile.mkdtemp(prefix="recipe-book-"))
    conversion_directory = None
    try:
        docx_path = temporary_root / "recipe-book.docx"
        build_recipe_book_docx(recipes, title, docx_path, include_notes)
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

    def show(self, initial_item_id=None):
        app = self.app
        window = tk.Toplevel(app.root)
        window.title("מרכז המתכונים")
        window.geometry("1160x800")
        window.minsize(920, 650)
        window.transient(app.root)
        window.configure(background=app.colors["background"])
        outer = ttk.Frame(window, padding=16, style="App.TFrame")
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="מרכז המתכונים", style="Section.TLabel").pack(fill="x", anchor="e")
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
            }
            self.records.append(record)
            self.by_id[record["id"]] = record

        if not self.records:
            ttk.Label(outer, text="הספרייה עדיין ריקה. הוסיפו קבצים או מתכונים כדי להתחיל.", style="Muted.TLabel").pack(fill="both", expand=True)
            ttk.Button(outer, text="סגירה", command=window.destroy).pack(anchor="w")
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
        ttk.Checkbutton(options, text="לכלול הערות אישיות", variable=notes_var).pack(anchor="e", pady=3)
        ttk.Checkbutton(options, text="לפתוח PDF לאחר היצירה", variable=open_var).pack(anchor="e", pady=3)

        def create_pdf():
            chosen = selected(print_box)
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
                result = export_recipe_book_pdf(chosen, title, Path(target), self.converter, notes_var.get())
                app.status_value.set(f"חוברת PDF נוצרה: {result.name}")
                messagebox.showinfo("החוברת מוכנה", f"קובץ PDF נוצר בהצלחה.\n\n{result}\n\nאפשר לפתוח אותו ולהדפיס במדפסת הרגילה.", parent=window)
                if open_var.get():
                    try:
                        os.startfile(str(result))
                    except (AttributeError, OSError):
                        webbrowser.open(result.as_uri())
            except Exception as exc:
                messagebox.showerror("שגיאה ביצירת PDF", f"לא ניתן ליצור את החוברת.\n{exc}\n\nבדקו שתיקיית LibreOffice נמצאת לצד קובץ התוכנה.", parent=window)

        ttk.Button(options, text="יצירת חוברת PDF  🖨", style="Primary.TButton", command=create_pdf).pack(fill="x", pady=(18, 6))
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

        # 4. Step-by-step instructions for cooking.
        cooking_tab = ttk.Frame(notebook, padding=16, style="App.TFrame")
        notebook.add(cooking_tab, text="מצב בישול")
        ttk.Label(cooking_tab, text="מצב הכנה — שלב אחר שלב", style="Section.TLabel").pack(fill="x", anchor="e")
        ttk.Label(cooking_tab, text="בחרו מתכון והתקדמו בין שלבי ההכנה.", style="Muted.TLabel").pack(fill="x", anchor="e", pady=(3, 12))
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
        step_label = ttk.Label(card, text="", style="Section.TLabel", justify="right", anchor="ne", wraplength=900)
        step_label.pack(fill="both", expand=True, anchor="ne")
        steps = []
        current = [0]

        def render_step():
            if not steps:
                count_label.configure(text="אין שלבי הכנה זמינים.")
                step_label.configure(text="אפשר להוסיף הוראות בחלון ״תצוגת מתכון מקצועית״.")
                return
            count_label.configure(text=f"שלב {current[0] + 1} מתוך {len(steps)}")
            step_label.configure(text=steps[current[0]])

        def load_cooking(_event=None):
            item_id = id_by_display.get(cooking_var.get())
            steps.clear()
            current[0] = 0
            if item_id is not None:
                recipe = self.by_id[item_id]
                raw = recipe.get("instructions") or recipe.get("content") or ""
                steps.extend([line.strip() for line in raw.splitlines() if line.strip()])
            render_step()

        cooking_combo.bind("<<ComboboxSelected>>", load_cooking)
        navigation = ttk.Frame(cooking_tab, style="App.TFrame")
        navigation.pack(fill="x")
        ttk.Button(navigation, text="השלב הקודם", command=lambda: (current.__setitem__(0, max(0, current[0] - 1)), render_step())).pack(side="right")
        ttk.Button(navigation, text="השלב הבא", style="Primary.TButton", command=lambda: (current.__setitem__(0, min(max(0, len(steps) - 1), current[0] + 1)), render_step())).pack(side="right", padx=8)
        if initial_item_id is not None and int(initial_item_id) in display_by_id:
            cooking_var.set(display_by_id[int(initial_item_id)])
            load_cooking()
        return window
