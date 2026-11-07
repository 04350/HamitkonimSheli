# -*- coding: utf-8 -*-
"""Offline OCR helpers for recipe pages.

The scanner keeps page geometry from Tesseract so title detection can prefer
large display text instead of blindly selecting the first text line.
"""
from __future__ import annotations

import re
import statistics
from typing import Any

try:
    from PIL import Image, ImageEnhance, ImageFilter, ImageOps
except Exception:  # pragma: no cover - optional dependency guard
    Image = ImageEnhance = ImageFilter = ImageOps = None


_HEBREW = re.compile(r"[א-ת]")
_LETTERS = re.compile(r"[A-Za-zא-ת]")
_HEADER_ONLY = {
    "מצרכים", "רכיבים", "החומרים", "רשימת מצרכים",
    "אופן ההכנה", "הוראות הכנה", "שלבי הכנה", "הכנה",
    "ingredients", "method", "directions", "preparation",
    "recipe", "מתכון",
}


def preprocess_image(image):
    """Return an enhanced grayscale image and a high-contrast variant."""
    if ImageOps is None:
        return image, image
    source = ImageOps.exif_transpose(image).convert("RGB")
    width, height = source.size
    shortest = min(width, height)
    if shortest and shortest < 1400:
        scale = min(2.2, 1400 / shortest)
        source = source.resize(
            (max(1, int(width * scale)), max(1, int(height * scale))),
            getattr(getattr(Image, "Resampling", Image), "LANCZOS", 1),
        )
    gray = ImageOps.grayscale(source)
    gray = ImageOps.autocontrast(gray, cutoff=1)
    if ImageFilter is not None:
        gray = gray.filter(ImageFilter.MedianFilter(size=3))
        gray = gray.filter(ImageFilter.UnsharpMask(radius=1.4, percent=145, threshold=3))
    enhanced = ImageEnhance.Contrast(gray).enhance(1.12) if ImageEnhance else gray
    threshold = enhanced.point(lambda value: 255 if value >= 178 else 0)
    return enhanced, threshold


def lines_from_tesseract_data(data: dict[str, Any]) -> list[dict[str, Any]]:
    """Rebuild lines from image_to_data output while retaining text geometry."""
    groups: dict[tuple[int, int, int], list[dict[str, Any]]] = {}
    texts = data.get("text", []) or []
    count = len(texts)
    for index in range(count):
        token = str(texts[index] or "").strip()
        if not token:
            continue
        try:
            confidence = float(data.get("conf", ["0"] * count)[index])
        except (TypeError, ValueError, IndexError):
            confidence = 0.0
        if confidence < 0:
            continue

        def number(name):
            try:
                return int(data.get(name, [0] * count)[index])
            except (TypeError, ValueError, IndexError):
                return 0

        key = (number("block_num"), number("par_num"), number("line_num"))
        groups.setdefault(key, []).append({
            "text": token,
            "confidence": confidence,
            "left": number("left"),
            "top": number("top"),
            "width": number("width"),
            "height": number("height"),
        })

    lines = []
    for words in groups.values():
        hebrew_chars = sum(len(_HEBREW.findall(word["text"])) for word in words)
        latin_chars = sum(len(re.findall(r"[A-Za-z]", word["text"])) for word in words)
        words.sort(key=lambda word: word["left"], reverse=hebrew_chars >= latin_chars)
        heights = [word["height"] for word in words if word["height"] > 0]
        if not heights:
            continue
        lines.append({
            "text": " ".join(word["text"] for word in words).strip(),
            "top": min(word["top"] for word in words),
            "left": min(word["left"] for word in words),
            "right": max(word["left"] + word["width"] for word in words),
            "height": int(statistics.median(heights)),
            "confidence": round(statistics.mean(word["confidence"] for word in words), 1),
        })
    return sorted(lines, key=lambda line: (line["top"], -line["right"]))


def select_title_line(lines: list[dict[str, Any]], page_height: int | None = None) -> str:
    """Choose a likely title using font/text height first, then confidence and position."""
    candidates = []
    useful_heights = []
    for line in lines:
        text = re.sub(r"\s+", " ", str(line.get("text", ""))).strip(" \t:：-•▪●")
        if not text or not _LETTERS.search(text):
            continue
        normalized = re.sub(r"[\W_]+", " ", text, flags=re.UNICODE).strip().casefold()
        if normalized in _HEADER_ONLY:
            continue
        if len(text) < 3 or len(text) > 150 or len(text.split()) > 20:
            continue
        if re.fullmatch(r"[\d\s.,:/-]+", text):
            continue
        top = max(0, int(line.get("top", 0) or 0))
        height = max(1, int(line.get("height", 1) or 1))
        if page_height and top > page_height * 0.72:
            continue
        useful_heights.append(height)
        candidates.append((line, text, top, height))
    if not candidates:
        return ""
    typical = statistics.median(useful_heights) if useful_heights else 1
    scored = []
    for line, text, top, height in candidates:
        confidence = max(0.0, min(100.0, float(line.get("confidence", 50) or 0)))
        upper_bonus = 1.18 if page_height and top < page_height * 0.32 else (
            1.05 if page_height and top < page_height * 0.58 else 0.82
        )
        size_ratio = height / max(1.0, typical)
        score = (size_ratio ** 1.55) * (0.78 + confidence / 250.0) * upper_bonus
        if len(text) <= 90:
            score *= 1.04
        scored.append((score, confidence, text))
    scored.sort(reverse=True)
    best = scored[0]
    return best[2] if best[1] >= 12 else ""


def best_title_from_pdf_page(page) -> str:
    """Use embedded PDF font sizes when a searchable text layer exists."""
    try:
        page_dict = page.get_text("dict")
        lines = []
        for block in page_dict.get("blocks", []):
            for line in block.get("lines", []):
                spans = [span for span in line.get("spans", []) if str(span.get("text", "")).strip()]
                if not spans:
                    continue
                text = " ".join(str(span.get("text", "")).strip() for span in spans).strip()
                size = statistics.mean(float(span.get("size", 0) or 0) for span in spans)
                bbox = line.get("bbox", (0, 0, 0, 0))
                lines.append({
                    "text": text,
                    "top": int(bbox[1]) if len(bbox) > 1 else 0,
                    "height": max(1, int(size * 2)),
                    "confidence": 95,
                })
        return select_title_line(lines, int(page.rect.height))
    except Exception:
        return ""


def recognize_image(image, pytesseract_module=None) -> dict[str, Any]:
    """OCR a page and return text, geometry-based title and confidence."""
    if pytesseract_module is None:
        try:
            import pytesseract as pytesseract_module
        except Exception:
            return {"text": "", "title": "", "lines": [], "confidence": 0.0}
    if image is None:
        return {"text": "", "title": "", "lines": [], "confidence": 0.0}

    try:
        enhanced, threshold = preprocess_image(image)
    except Exception:
        enhanced = image
        threshold = image
    variants = [enhanced]
    if threshold is not enhanced:
        variants.append(threshold)

    attempts = []
    for variant in variants:
        for psm in (6, 11):
            try:
                data = pytesseract_module.image_to_data(
                    variant,
                    lang="heb+eng",
                    config=f"--oem 3 --psm {psm}",
                    output_type=pytesseract_module.Output.DICT,
                    timeout=60,
                )
                lines = lines_from_tesseract_data(data)
                if not lines:
                    continue
                full_text = "\n".join(line["text"] for line in lines)
                confidences = [line["confidence"] for line in lines if line["confidence"] > 0]
                average_confidence = statistics.mean(confidences) if confidences else 0.0
                title = select_title_line(lines, getattr(variant, "height", 0))
                quality = min(1000, len(re.sub(r"\s+", "", full_text))) + average_confidence * 2
                attempts.append({
                    "quality": quality,
                    "text": full_text,
                    "title": title,
                    "lines": lines,
                    "confidence": round(average_confidence, 1),
                })
            except Exception:
                continue

    if attempts:
        attempts.sort(key=lambda item: item["quality"], reverse=True)
        best = attempts[0]
        title_candidates = [item for item in attempts if item.get("title")]
        if title_candidates:
            title_candidates.sort(
                key=lambda item: max(
                    (line.get("height", 0) * (0.8 + line.get("confidence", 0) / 250)
                     for line in item["lines"] if line.get("text") == item["title"]),
                    default=0,
                ),
                reverse=True,
            )
            best["title"] = title_candidates[0]["title"]
        return {key: best[key] for key in ("text", "title", "lines", "confidence")}

    try:
        text = pytesseract_module.image_to_string(
            enhanced, lang="heb+eng", config="--oem 3 --psm 6", timeout=60
        )
        return {"text": text or "", "title": "", "lines": [], "confidence": 0.0}
    except Exception:
        return {"text": "", "title": "", "lines": [], "confidence": 0.0}


def best_title_from_image(image, pytesseract_module=None) -> str:
    return str(recognize_image(image, pytesseract_module).get("title") or "")
