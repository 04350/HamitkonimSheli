# -*- coding: utf-8 -*-
"""Resource-aware offline OCR helpers for recipe pages."""
from __future__ import annotations

import math
import re
import statistics
from typing import Any

try:
    from PIL import Image, ImageEnhance, ImageFilter, ImageOps
except Exception:  # pragma: no cover
    Image = ImageEnhance = ImageFilter = ImageOps = None


MAX_OCR_EDGE = 2800
MAX_OCR_PIXELS = 8_000_000
_HEBREW = re.compile(r"[א-ת]")
_LETTERS = re.compile(r"[A-Za-zא-ת]")
_HEADER_ONLY = {
    "מצרכים", "רכיבים", "החומרים", "רשימת מצרכים",
    "אופן ההכנה", "הוראות הכנה", "שלבי הכנה", "הכנה",
    "ingredients", "method", "directions", "preparation",
    "recipe", "מתכון",
}


def preprocess_image(image):
    """Downsample oversized scans before RGB conversion to cap OCR memory use."""
    if ImageOps is None:
        return image, image
    oriented = ImageOps.exif_transpose(image)
    width, height = oriented.size
    pixels = max(1, width * height)
    scale = min(1.0, MAX_OCR_EDGE / max(width, height, 1), math.sqrt(MAX_OCR_PIXELS / pixels))
    if scale < 1.0:
        size = (max(1, int(width * scale)), max(1, int(height * scale)))
        source = oriented.resize(size, getattr(getattr(Image, "Resampling", Image), "LANCZOS", 1))
    else:
        source = oriented.copy()
    if oriented is not image:
        try:
            oriented.close()
        except Exception:
            pass
    source = source.convert("RGB")
    gray = ImageOps.grayscale(source)
    source.close()
    adjusted = ImageOps.autocontrast(gray, cutoff=1)
    gray.close()
    if ImageFilter is not None:
        filtered = adjusted.filter(ImageFilter.MedianFilter(size=3))
        adjusted.close()
        sharpened = filtered.filter(ImageFilter.UnsharpMask(radius=1.2, percent=125, threshold=3))
        filtered.close()
        adjusted = sharpened
    enhanced = ImageEnhance.Contrast(adjusted).enhance(1.1) if ImageEnhance else adjusted.copy()
    threshold = enhanced.point(lambda value: 255 if value >= 178 else 0)
    adjusted.close()
    return enhanced, threshold


def lines_from_tesseract_data(data: dict[str, Any]) -> list[dict[str, Any]]:
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
            "text": token, "confidence": confidence, "left": number("left"),
            "top": number("top"), "width": number("width"), "height": number("height"),
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
    candidates = []
    useful_heights = []
    for line in lines:
        text = re.sub(r"\s+", " ", str(line.get("text", ""))).strip(" \t:：-•▪●")
        if not text or not _LETTERS.search(text):
            continue
        normalized = re.sub(r"[\W_]+", " ", text, flags=re.UNICODE).strip().casefold()
        if normalized in _HEADER_ONLY or len(text) < 3 or len(text) > 150 or len(text.split()) > 20:
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
        score = (height / max(1.0, typical)) ** 1.55 * (0.78 + confidence / 250.0) * upper_bonus
        if len(text) <= 90:
            score *= 1.04
        scored.append((score, confidence, text))
    scored.sort(reverse=True)
    best = scored[0]
    return best[2] if best[1] >= 12 else ""


def best_title_from_pdf_page(page) -> str:
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
                    "text": text, "top": int(bbox[1]) if len(bbox) > 1 else 0,
                    "height": max(1, int(size * 2)), "confidence": 95,
                })
        return select_title_line(lines, int(page.rect.height))
    except Exception:
        return ""


def _recognize_attempt(image, pytesseract_module, psm: int) -> dict[str, Any] | None:
    data = pytesseract_module.image_to_data(
        image, lang="heb+eng", config=f"--oem 3 --psm {psm}",
        output_type=pytesseract_module.Output.DICT, timeout=35,
    )
    lines = lines_from_tesseract_data(data)
    if not lines:
        return None
    full_text = "\n".join(line["text"] for line in lines)
    confidences = [line["confidence"] for line in lines if line["confidence"] > 0]
    average_confidence = statistics.mean(confidences) if confidences else 0.0
    title = select_title_line(lines, getattr(image, "height", 0))
    meaningful_length = len(re.sub(r"\s+", "", full_text))
    quality = min(2000, meaningful_length * 2) + average_confidence * 3 + (80 if title else 0)
    return {
        "quality": quality, "text": full_text, "title": title,
        "lines": lines, "confidence": round(average_confidence, 1),
    }


def recognize_image(image, pytesseract_module=None) -> dict[str, Any]:
    """Use one normal OCR pass, retry with alternate layouts only when confidence is poor."""
    if pytesseract_module is None:
        try:
            import pytesseract as pytesseract_module
        except Exception:
            return {"text": "", "title": "", "lines": [], "confidence": 0.0}
    if image is None:
        return {"text": "", "title": "", "lines": [], "confidence": 0.0}

    enhanced = threshold = None
    attempts = []
    try:
        enhanced, threshold = preprocess_image(image)
        first = _recognize_attempt(enhanced, pytesseract_module, 6)
        if first:
            attempts.append(first)
        retry_needed = (
            first is None
            or len(re.sub(r"\s+", "", first["text"])) < 12
            or first["confidence"] < 30
            or not first["title"]
        )
        if retry_needed:
            second = _recognize_attempt(enhanced, pytesseract_module, 11)
            if second:
                attempts.append(second)
        if retry_needed and (not attempts or max(item["quality"] for item in attempts) < 120):
            third = _recognize_attempt(threshold, pytesseract_module, 6)
            if third:
                attempts.append(third)
        if attempts:
            best = max(attempts, key=lambda item: item["quality"])
            return {key: best[key] for key in ("text", "title", "lines", "confidence")}
        fallback = pytesseract_module.image_to_string(
            enhanced, lang="heb+eng", config="--oem 3 --psm 6", timeout=35
        )
        return {"text": fallback or "", "title": "", "lines": [], "confidence": 0.0}
    except Exception:
        return {"text": "", "title": "", "lines": [], "confidence": 0.0}
    finally:
        for candidate in (enhanced, threshold):
            if candidate is not None:
                try:
                    candidate.close()
                except Exception:
                    pass


def best_title_from_image(image, pytesseract_module=None) -> str:
    return str(recognize_image(image, pytesseract_module).get("title") or "")
