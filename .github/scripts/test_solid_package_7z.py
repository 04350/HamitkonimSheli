# -*- coding: utf-8 -*-
"""Measure a one-file solid 7z distribution without removing any runtime files."""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
import tempfile
import time
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from py7zr import FILTER_LZMA2, SevenZipFile
from docx import Document
import openpyxl
from pptx import Presentation

TARGET_BYTES = 300_000_000
PACKAGE_ROOT = Path("source-artifact")
WORK = Path("solid-7z-experiment").resolve()
WORK.mkdir(parents=True, exist_ok=True)

packages = list(PACKAGE_ROOT.rglob("*.zip"))
if len(packages) != 1:
    raise SystemExit(f"Expected one source package ZIP, found {packages}")
source_package = packages[0]

with zipfile.ZipFile(source_package, "r") as archive:
    if archive.testzip() is not None:
        raise SystemExit("Original distribution ZIP failed CRC validation")
    names = archive.namelist()
    exe_name = next((n for n in names if n.lower().endswith(".exe")), None)
    office_name = next((n for n in names if Path(n).name.lower() == "libreoffice.zip"), None)
    if not exe_name or not office_name:
        raise SystemExit(f"Missing EXE or LibreOffice.zip: {names}")
    exe_path = WORK / Path(exe_name).name
    office_zip = WORK / "LibreOffice-source.zip"
    with archive.open(exe_name) as source, exe_path.open("wb") as target:
        shutil.copyfileobj(source, target, 4 * 1024 * 1024)
    with archive.open(office_name) as source, office_zip.open("wb") as target:
        shutil.copyfileobj(source, target, 4 * 1024 * 1024)

office_root = WORK / "LibreOffice"
shutil.rmtree(office_root, ignore_errors=True)
office_root.mkdir(parents=True)
with zipfile.ZipFile(office_zip, "r") as archive:
    if archive.testzip() is not None:
        raise SystemExit("Original LibreOffice ZIP failed CRC validation")
    archive.extractall(office_root)
    original_names = sorted(i.filename for i in archive.infolist() if not i.is_dir())
    original_sizes = {i.filename: i.file_size for i in archive.infolist() if not i.is_dir()}
print(f"Original app EXE: {exe_path.stat().st_size:,} bytes", flush=True)
print(f"LibreOffice files: {len(original_names):,}", flush=True)
print(f"LibreOffice expanded size: {sum(original_sizes.values()):,} bytes", flush=True)

# Prepare test documents once. All conversions below use the same
# find_libreoffice() and convert_office_to_pdf() functions shipped in the app.
sample_dir = WORK / "sample-documents"
sample_dir.mkdir(exist_ok=True)

docx_path = sample_dir / "בדיקת-Word.docx"
doc = Document()
doc.add_heading("בדיקת חבילת Office", level=1)
doc.add_paragraph("בדיקת תקינות המרה למסמך PDF.")
doc.save(docx_path)

xlsx_path = sample_dir / "בדיקת-Excel.xlsx"
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet["A1"] = "בדיקת Excel"
worksheet["A2"] = 123
workbook.save(xlsx_path)

pptx_path = sample_dir / "בדיקת-PowerPoint.pptx"
presentation = Presentation()
slide = presentation.slides.add_slide(presentation.slide_layouts[6])
textbox = slide.shapes.add_textbox(500000, 500000, 7000000, 1200000)
textbox.text = "Office conversion check"
presentation.save(pptx_path)
samples = [docx_path, xlsx_path, pptx_path]

sys.path.insert(0, str(Path("src").resolve()))
import ספריית_המתכונים as app

old_executable = sys.executable
had_frozen = hasattr(sys, "frozen")
old_frozen = getattr(sys, "frozen", None)

def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            block = handle.read(4 * 1024 * 1024)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()

def verify_identical_tree(source_root: Path, target_root: Path) -> dict:
    source_files = {p.relative_to(source_root).as_posix(): p for p in source_root.rglob("*") if p.is_file()}
    target_files = {p.relative_to(target_root).as_posix(): p for p in target_root.rglob("*") if p.is_file()}
    if set(source_files) != set(target_files):
        missing = sorted(set(source_files) - set(target_files))[:20]
        extra = sorted(set(target_files) - set(source_files))[:20]
        raise RuntimeError(f"Extracted file list differs; missing={missing}; extra={extra}")
    checked = 0
    for relative, source in source_files.items():
        target = target_files[relative]
        if source.stat().st_size != target.stat().st_size:
            raise RuntimeError(f"Size differs after 7z extraction: {relative}")
        if sha256_file(source) != sha256_file(target):
            raise RuntimeError(f"SHA256 differs after 7z extraction: {relative}")
        checked += 1
    return {"file_count": checked, "all_file_bytes_identical": True}

def run_real_app_checks(exe_file: Path, lo_folder: Path, label: str) -> list[str]:
    local_data = WORK / f"local-data-{label}"
    shutil.rmtree(local_data, ignore_errors=True)
    os.environ["LOCALAPPDATA"] = str(local_data)
    sys.executable = str(exe_file.resolve())
    sys.frozen = True
    engine = app.find_libreoffice()
    expected_engine = lo_folder / "program" / "soffice.exe"
    if engine is None or engine.parent.name.lower() != "program" or engine.name.lower() not in ("soffice.com", "soffice.exe"):
        raise RuntimeError(f"{label}: the application did not discover the extracted LibreOffice folder: {engine}")
    expected_engine = lo_folder / "program" / engine.name
    if engine.resolve() != expected_engine.resolve():
        raise RuntimeError(f"{label}: LibreOffice was discovered outside the extracted runtime folder: {engine}")
    converted = []
    for source in samples:
        pdf, temp_dir, error = app.convert_office_to_pdf(source)
        if error or pdf is None or not pdf.is_file():
            raise RuntimeError(f"{label}: failed to convert {source.suffix}: {error}")
        try:
            with app.fitz.open(pdf) as result:
                if result.page_count < 1:
                    raise RuntimeError(f"{label}: {source.suffix} produced an empty PDF")
                converted.append(f"{source.suffix.upper()}={result.page_count}p")
        finally:
            if temp_dir:
                shutil.rmtree(temp_dir, ignore_errors=True)
    return converted

variants = [
    ("LZMA2-solid-preset9", [{"id": FILTER_LZMA2, "preset": 9}]),
]

results = []
try:
    for label, filters in variants:
        print(f"\n=== Building {label} ===", flush=True)
        candidate = WORK / f"recipe-library-{label}.7z"
        started = time.time()
        with SevenZipFile(candidate, mode="w", filters=filters) as archive:
            archive.write(exe_path, arcname=exe_path.name)
            archive.writeall(office_root, arcname="LibreOffice")
        compress_seconds = round(time.time() - started, 1)
        size = candidate.stat().st_size
        print(f"7z candidate size: {size:,} bytes ({size / 1_000_000:.2f} MB); target <= {TARGET_BYTES:,} bytes; under_target={size <= TARGET_BYTES}", flush=True)
        with SevenZipFile(candidate, mode="r") as archive:
            info = archive.archiveinfo()
            if archive.testzip() is not None:
                raise RuntimeError(f"{label}: 7z archive CRC validation failed")
            is_solid = bool(info.solid)
            names = archive.getnames()
        if not is_solid:
            raise RuntimeError(f"{label}: expected solid 7z archive but archiveinfo().solid is false")

        extracted = WORK / f"extracted-{label}"
        shutil.rmtree(extracted, ignore_errors=True)
        extracted.mkdir(parents=True)
        with SevenZipFile(candidate, mode="r") as archive:
            archive.extractall(path=extracted)
        extracted_exe = extracted / exe_path.name
        extracted_lo = extracted / "LibreOffice"
        if not extracted_exe.is_file():
            raise RuntimeError(f"{label}: EXE missing after extraction")
        identical = verify_identical_tree(office_root, extracted_lo)
        conversions = run_real_app_checks(extracted_exe, extracted_lo, label)
        result = {
            "method": label,
            "archive_bytes": size,
            "archive_decimal_mb": round(size / 1_000_000, 2),
            "under_300mb": size <= TARGET_BYTES,
            "solid": is_solid,
            "file_entries": len(names),
            "compression_seconds": compress_seconds,
            "expanded_libreoffice_files": identical["file_count"],
            "all_libreoffice_file_bytes_identical": identical["all_file_bytes_identical"],
            "office_conversions": conversions,
        }
        results.append(result)
        print(json.dumps(result, ensure_ascii=False), flush=True)
finally:
    sys.executable = old_executable
    if had_frozen:
        sys.frozen = old_frozen
    elif hasattr(sys, "frozen"):
        del sys.frozen

report = {
    "source_package_bytes": source_package.stat().st_size,
    "target_bytes": TARGET_BYTES,
    "variants": results,
    "features_or_office_files_removed": False,
}
Path("solid-7z-results.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
lines = [
    "## תוצאות בדיקת 7z רציף",
    "",
    f'- גודל חבילה מקורית: **{source_package.stat().st_size / 1_000_000:.2f} MB**',
    "- נכללו כל הקבצים המקוריים, ללא הסרת רכיבים.",
    "",
    "| שיטה | גודל 7z | מתחת ל־300MB | Solid | קבצי Office זהים | המרות |",
    "|---|---:|:---:|:---:|:---:|---|",
]
for x in results:
    lines.append(
        f'| {x["method"]} | {x["archive_decimal_mb"]:.2f} MB | '
        f'{"כן" if x["under_300mb"] else "לא"} | כן | כן ({x["expanded_libreoffice_files"]:,}) | '
        f'{", ".join(x["office_conversions"])} |'
    )
with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as handle:
    handle.write("\n".join(lines) + "\n")
