# -*- coding: utf-8 -*-
"""Create disposable sample recipes for a real Windows UI screenshot in CI."""
from __future__ import annotations

import os
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import ספריית_המתכונים as app  # noqa: E402


RECIPES = [
    {
        "title": "לזניה ביתית עשירה",
        "category": "ארוחות עיקריות",
        "prep": "25", "cook": "55", "servings": "8", "rating": "5",
        "ingredients": "עלי לזניה\n500 גרם בשר טחון\nרוטב עגבניות\nגבינת מוצרלה\nבצל ושום",
        "instructions": "מטגנים בצל ושום. מוסיפים בשר ורוטב.\nמסדרים שכבות לזניה, רוטב וגבינה.\nאופים עד שהגבינה מזהיבה.",
        "colors": ((242, 188, 128), (189, 81, 44), (68, 118, 66)),
    },
    {
        "title": "עוגת שוקולד עשירה",
        "category": "קינוחים",
        "prep": "20", "cook": "40", "servings": "12", "rating": "5",
        "ingredients": "2 כוסות קמח\nכוס סוכר\n3 ביצים\n200 גרם שוקולד\nכוס שמן",
        "instructions": "מערבבים את החומרים היבשים.\nמוסיפים ביצים, שמן ושוקולד מומס.\nאופים כ-40 דקות.",
        "colors": ((112, 62, 42), (70, 38, 31), (151, 58, 62)),
    },
    {
        "title": "סלט ירקות צבעוני",
        "category": "סלטים",
        "prep": "15", "cook": "0", "servings": "5", "rating": "4",
        "ingredients": "מלפפונים\nעגבניות שרי\nפלפל אדום\nעלי חסה\nשמן זית ולימון",
        "instructions": "חותכים את הירקות לקוביות.\nמתבלים בשמן זית, לימון ומעט מלח.\nמערבבים ומגישים.",
        "colors": ((77, 142, 81), (224, 103, 64), (242, 192, 76)),
    },
    {
        "title": "מרק ירקות ביתי",
        "category": "מרקים",
        "prep": "20", "cook": "45", "servings": "6", "rating": "5",
        "ingredients": "גזר\nתפוחי אדמה\nקישוא\nבצל\nפטרוזיליה",
        "instructions": "חותכים את הירקות.\nמבשלים בסיר עם מים ותבלינים.\nמגישים חם עם עשבי תיבול.",
        "colors": ((224, 170, 70), (180, 95, 38), (85, 130, 66)),
    },
    {
        "title": "פסטה ברוטב שמנת",
        "category": "ארוחות עיקריות",
        "prep": "10", "cook": "20", "servings": "4", "rating": "4",
        "ingredients": "חבילת פסטה\nשמנת לבישול\nפטריות\nשום\nפרמזן",
        "instructions": "מבשלים את הפסטה.\nמטגנים פטריות ושום ומוסיפים שמנת.\nמערבבים עם הפסטה ומגישים.",
        "colors": ((245, 215, 148), (202, 158, 91), (105, 133, 69)),
    },
    {
        "title": "קציצות עוף עסיסיות",
        "category": "בשר ועוף",
        "prep": "20", "cook": "30", "servings": "6", "rating": "4",
        "ingredients": "500 גרם עוף טחון\nבצל מגורד\nפירורי לחם\nביצה\nפטרוזיליה",
        "instructions": "מערבבים את כל המרכיבים.\nיוצרים קציצות ומטגנים או אופים.\nמגישים עם תוספת.",
        "colors": ((192, 118, 72), (141, 74, 43), (76, 121, 65)),
    },
]


def create_cover(path: Path, index: int, dish: dict):
    width, height = 640, 420
    image = Image.new("RGB", (width, height))
    pixels = image.load()
    top, bottom, garnish = dish["colors"]
    for y in range(height):
        blend = y / max(1, height - 1)
        color = tuple(int(top[k] * (1 - blend) + bottom[k] * blend) for k in range(3))
        for x in range(width):
            vignette = max(0.82, 1 - abs(x - width / 2) / 1500)
            pixels[x, y] = tuple(int(channel * vignette) for channel in color)

    draw = ImageDraw.Draw(image)
    # A stylized plate, food pieces and herbs: decorative demo artwork only.
    draw.ellipse((72, 25, 568, 410), fill=(30, 34, 29))
    draw.ellipse((88, 14, 552, 398), fill=(248, 246, 235), outline=(221, 216, 202), width=8)
    draw.ellipse((122, 48, 518, 365), fill=(236, 229, 210))
    food_colors = [bottom, top, garnish, (226, 185, 117), (165, 58, 51), (95, 118, 60)]
    for piece in range(14):
        x = 175 + ((piece * 83 + index * 41) % 290)
        y = 95 + ((piece * 59 + index * 31) % 220)
        color = food_colors[(piece + index) % len(food_colors)]
        rx = 24 + (piece % 3) * 7
        ry = 14 + (piece % 4) * 5
        draw.ellipse((x - rx, y - ry, x + rx, y + ry), fill=color, outline=(245, 232, 199), width=2)
    for leaf in range(7):
        x = 250 + ((leaf * 41 + index * 22) % 170)
        y = 120 + ((leaf * 29 + index * 13) % 150)
        draw.ellipse((x - 18, y - 8, x + 18, y + 8), fill=garnish)
    image = image.filter(ImageFilter.GaussianBlur(0.25))
    image.save(path, quality=91)


def main():
    # Windows runners may default to cp1252; demo paths contain Hebrew.
    # Use UTF-8 explicitly so successful demo creation cannot fail on logging.
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    profile = os.environ.get("LOCALAPPDATA")
    if not profile:
        raise SystemExit("LOCALAPPDATA must point to an isolated demo profile.")
    library = app.RecipeLibrary()
    photos_dir = library.files_dir / "תמונות מתכונים"
    photos_dir.mkdir(parents=True, exist_ok=True)
    sources_dir = Path(profile) / "RecipeDemoSources"
    sources_dir.mkdir(parents=True, exist_ok=True)

    for index, recipe in enumerate(RECIPES):
        source = sources_dir / f"demo-{index + 1}.txt"
        source.write_text(
            f"{recipe['title']}\n\nמצרכים:\n{recipe['ingredients']}\n\nאופן ההכנה:\n{recipe['instructions']}\n",
            encoding="utf-8"
        )
        prepared = library.prepare_import(source)
        added, title = library.import_prepared(prepared)
        if not added:
            continue
        row = library.db.execute(
            "SELECT id FROM items WHERE digest=?", (prepared["digest"],)
        ).fetchone()
        if not row:
            continue
        item_id = int(row["id"])
        image_path = photos_dir / f"demo-{index + 1}.jpg"
        create_cover(image_path, index, recipe)
        data = library.get_recipe_data(item_id)
        data.update({
            "category": recipe["category"],
            "ingredients": recipe["ingredients"],
            "instructions": recipe["instructions"],
            "prep_minutes": recipe["prep"],
            "cook_minutes": recipe["cook"],
            "servings": recipe["servings"],
            "rating": recipe["rating"],
            "is_recipe": True,
            "image_path": str(image_path),
        })
        library.set_recipe_data(item_id, data)
        if index in (0, 1):
            library.set_favorite(item_id, True)

    library.set_setting("color_theme", "מרווה")
    library.set_setting("font_size", "10")
    library.set_setting("reader_preview_mode", "image")
    total_rows = library.db.execute("SELECT COUNT(*) FROM items WHERE is_deleted=0").fetchone()[0]
    recipe_rows = [
        row for row in library.rows()
        if app.is_recipe_record(library.get_recipe_data(row["id"]), row["content"] or "")
    ]
    if total_rows < len(RECIPES) or len(recipe_rows) < len(RECIPES):
        library.db.close()
        raise RuntimeError(
            f"Demo data incomplete: items={total_rows}; detected_recipes={len(recipe_rows)}; expected={len(RECIPES)}"
        )
    print(f"Demo recipe library ready: items={total_rows}; recipes={len(recipe_rows)}; data_dir={library.home}")
    library.db.close()


if __name__ == "__main__":
    main()
