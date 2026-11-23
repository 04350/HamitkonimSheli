#!/usr/bin/env python3
"""Validate Tesseract PNG/TIFF line images and matching single-line .gt.txt files."""
from __future__ import annotations
import argparse,json,sys
from pathlib import Path
IMAGE_SUFFIXES={".png",".tif",".tiff"}
def inspect_dataset(directory:Path)->dict:
    if not directory.is_dir(): raise ValueError(f"Dataset directory does not exist: {directory}")
    images=sorted(p for p in directory.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES)
    texts=sorted(directory.glob("*.gt.txt")); image_stems={p.stem for p in images}
    text_stems={p.name[:-len(".gt.txt")] for p in texts}
    report={"directory":str(directory),"image_count":len(images),"transcription_count":len(texts),
      "valid_samples":0,"missing_transcriptions":sorted(image_stems-text_stems),
      "missing_images":sorted(text_stems-image_stems),"empty_transcriptions":[],
      "multiline_transcriptions":[],"invalid_utf8":[]}
    for image in images:
        gt=directory/(image.stem+".gt.txt")
        if not gt.is_file():continue
        try:value=gt.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:report["invalid_utf8"].append(gt.name);continue
        value=value.strip()
        if not value:report["empty_transcriptions"].append(gt.name)
        elif "\n" in value or "\r" in value:report["multiline_transcriptions"].append(gt.name)
        else:report["valid_samples"]+=1
    return report
def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__);p.add_argument("directory",type=Path)
    p.add_argument("--min-samples",type=int,default=1);p.add_argument("--json",dest="json_path",type=Path)
    a=p.parse_args(argv)
    if a.min_samples<1:p.error("--min-samples must be at least 1")
    try:r=inspect_dataset(a.directory)
    except ValueError as e:print(f"ERROR: {e}",file=sys.stderr);return 2
    rendered=json.dumps(r,ensure_ascii=False,indent=2);print(rendered)
    if a.json_path:
        a.json_path.parent.mkdir(parents=True,exist_ok=True);a.json_path.write_text(rendered+"\n",encoding="utf-8")
    invalid=any(r[k] for k in ("missing_transcriptions","missing_images","empty_transcriptions","multiline_transcriptions","invalid_utf8"))
    return 1 if invalid or r["valid_samples"]<a.min_samples else 0
if __name__=="__main__":raise SystemExit(main())
