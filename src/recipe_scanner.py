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


def _otsu_threshold(image) -> int:
    """Choose a per-image threshold from its grayscale histogram (Otsu)."""
    histogram = image.histogram()[:256]
    total = sum(histogram)
    if not total:
        return 178
    weighted_total = sum(level * count for level, count in enumerate(histogram))
    background_weight = 0
    background_sum = 0
    best_variance = -1.0
    best_threshold = 178
    for level, count in enumerate(histogram):
        background_weight += count
        if not background_weight:
            continue
        foreground_weight = total - background_weight
        if not foreground_weight:
            break
        background_sum += level * count
        background_mean = background_sum / background_weight
        foreground_mean = (weighted_total - background_sum) / foreground_weight
        variance = background_weight * foreground_weight * (background_mean - foreground_mean) ** 2
        if variance > best_variance:
            best_variance = variance
            best_threshold = level
    return max(105, min(225, best_threshold))


def _crop_page_margins(gray):
    """Remove only large, clearly blank page margins; preserve the page otherwise."""
    width, height = gray.size
    if width < 100 or height < 100:
        return gray.copy()
    threshold = max(185, min(238, _otsu_threshold(gray) + 25))
    mask = gray.point(lambda value: 255 if value < threshold else 0)
    try:
        bounds = mask.getbbox()
    finally:
        mask.close()
    if not bounds:
        return gray.copy()
    left, top, right, bottom = bounds
    pad_x = max(10, int(width * 0.025))
    pad_y = max(10, int(height * 0.025))
    left = max(0, left - pad_x)
    top = max(0, top - pad_y)
    right = min(width, right + pad_x)
    bottom = min(height, bottom + pad_y)
    crop_width, crop_height = right - left, bottom - top
    # Avoid cropping photographs, colored paper, or pages with content near an edge.
    if crop_width < width * 0.30 or crop_height < height * 0.30:
        return gray.copy()
    retained = (crop_width * crop_height) / max(1, width * height)
    if retained > 0.92:
        return gray.copy()
    return gray.crop((left, top, right, bottom))


def preprocess_image(image):
    """Normalize orientation, safely trim blank margins, and create adaptive OCR variants."""
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
    cropped = _crop_page_margins(gray)
    gray.close()
    adjusted = ImageOps.autocontrast(cropped, cutoff=1)
    cropped.close()
    if ImageFilter is not None:
        filtered = adjusted.filter(ImageFilter.MedianFilter(size=3))
        adjusted.close()
        sharpened = filtered.filter(ImageFilter.UnsharpMask(radius=1.2, percent=125, threshold=3))
        filtered.close()
        adjusted = sharpened
    enhanced = ImageEnhance.Contrast(adjusted).enhance(1.1) if ImageEnhance else adjusted.copy()
    binary_threshold = _otsu_threshold(enhanced)
    threshold = enhanced.point(lambda value: 255 if value >= binary_threshold else 0)
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
    """Recognize a scan and try alternate layouts/orientations only when the first result is weak."""
    empty = {"text": "", "title": "", "lines": [], "confidence": 0.0}
    if pytesseract_module is None:
        try:
            import pytesseract as pytesseract_module
        except Exception:
            return empty
    if image is None:
        return empty

    enhanced = threshold = None
    attempts = []
    rotated_images = []
    try:
        enhanced, threshold = preprocess_image(image)
        first = _recognize_attempt(enhanced, pytesseract_module, 6)
        if first:
            attempts.append(first)
        first_text_length = len(re.sub(r"\s+", "", first["text"])) if first else 0
        retry_needed = (
            first is None
            or (first_text_length < 8 and first["confidence"] < 60)
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

        best = max(attempts, key=lambda item: item["quality"]) if attempts else None
        best_text_length = len(re.sub(r"\s+", "", best["text"])) if best else 0
        # Rotated photos/scans are expensive to retry, so test them only after weak upright OCR.
        if best is None or best["confidence"] < 42 or best_text_length < 8:
            for angle in (90, 180, 270):
                rotated = enhanced.rotate(angle, expand=True)
                rotated_images.append(rotated)
                candidate = _recognize_attempt(rotated, pytesseract_module, 6)
                if candidate:
                    attempts.append(candidate)
                    best = max(attempts, key=lambda item: item["quality"])
                    best_text_length = len(re.sub(r"\s+", "", best["text"]))
                    if best["confidence"] >= 62 and best_text_length >= 8 and best["title"]:
                        break

        if attempts:
            best = max(attempts, key=lambda item: item["quality"])
            return {key: best[key] for key in ("text", "title", "lines", "confidence")}
        fallback = pytesseract_module.image_to_string(
            enhanced, lang="heb+eng", config="--oem 3 --psm 6", timeout=35
        )
        return {"text": fallback or "", "title": "", "lines": [], "confidence": 0.0}
    except Exception:
        return empty
    finally:
        for candidate in (enhanced, threshold, *rotated_images):
            if candidate is not None:
                try:
                    candidate.close()
                except Exception:
                    pass

def best_title_from_image(image, pytesseract_module=None) -> str:
    return str(recognize_image(image, pytesseract_module).get("title") or "")
