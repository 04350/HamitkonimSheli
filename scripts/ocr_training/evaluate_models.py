#!/usr/bin/env python3
"""Compare OCR models on manually transcribed recipe scans held out from training."""
from __future__ import annotations
import argparse,csv,json,re,subprocess,sys
from pathlib import Path
IMAGE_SUFFIXES={".png",".tif",".tiff"}
DEFAULT_MODELS="heb+eng,heb4+eng,recipe_heb+eng,recipe_heb4+eng"
def distance(a,b):
    row=list(range(len(b)+1))
    for i,ca in enumerate(a,1):
        cur=[i]
        for j,cb in enumerate(b,1):cur.append(min(cur[-1]+1,row[j]+1,row[j-1]+(ca!=cb)))
        row=cur
    return row[-1]
def norm(s):return re.sub(r"\s+"," ",s.replace("\ufeff","").strip())
def run_ocr(exe,image,tessdata,language,psm):
    cmd=[exe,str(image),"stdout","--tessdata-dir",str(tessdata),"-l",language,"--oem","1","--psm",str(psm)]
    done=subprocess.run(cmd,capture_output=True,text=True,encoding="utf-8",errors="replace",timeout=60,check=False)
    if done.returncode:raise RuntimeError(f"{language} failed on {image.name}: {done.stderr.strip()}")
    return done.stdout
def evaluate(eval_dir:Path,tessdata:Path,models:list[str],executable="tesseract",psm=7):
    images=sorted(p for p in eval_dir.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES)
    if not images:raise ValueError(f"No line images found in {eval_dir}")
    for im in images:
        gt=im.with_name(im.stem+".gt.txt")
        if not gt.is_file():raise ValueError(f"Missing transcription: {gt.name}")
        val=norm(gt.read_text(encoding="utf-8-sig"))
        if not val or "\n" in val or "\r" in val:raise ValueError(f"Invalid one-line transcription: {gt.name}")
    rows=[]; summaries=[]
    for model in models:
        missing=[x for x in model.split("+") if not (tessdata/f"{x}.traineddata").is_file()]
        if missing:print(f"Skipping {model}; missing model files: {missing}",file=sys.stderr);continue
        ce=we=de=chars=words=digits=exact=0
        for im in images:
            expected=norm(im.with_name(im.stem+".gt.txt").read_text(encoding="utf-8-sig"))
            predicted=norm(run_ocr(executable,im,tessdata,model,psm))
            c=distance(expected,predicted);w=distance(expected.split(),predicted.split())
            ed="".join(re.findall(r"\d",expected));pd="".join(re.findall(r"\d",predicted));d=distance(ed,pd)
            ce+=c;we+=w;de+=d;chars+=len(expected);words+=len(expected.split());digits+=len(ed);exact+=expected==predicted
            rows.append({"model":model,"image":im.name,"expected":expected,"predicted":predicted,
                         "cer":c/max(1,len(expected)),"wer":w/max(1,len(expected.split()))})
        summaries.append({"model":model,"samples":len(images),"cer":ce/max(1,chars),
          "wer":we/max(1,words),"digit_error_rate":de/max(1,digits) if digits else 0.0,
          "exact_line_accuracy":exact/max(1,len(images))})
    summaries.sort(key=lambda x:(x["cer"],x["digit_error_rate"],x["wer"]))
    return rows,{"evaluation_directory":str(eval_dir),"tessdata_directory":str(tessdata),
      "sample_count":len(images),"ranking":summaries,
      "best_model_by_cer_then_digits_then_wer":summaries[0]["model"] if summaries else None}
def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__);p.add_argument("--evaluation-dir",type=Path,required=True)
    p.add_argument("--tessdata-dir",type=Path,required=True);p.add_argument("--models",default=DEFAULT_MODELS)
    p.add_argument("--tesseract",default="tesseract");p.add_argument("--psm",type=int,default=7)
    p.add_argument("--output-dir",type=Path,default=Path("ocr-evaluation-results"));a=p.parse_args(argv)
    try:rows,report=evaluate(a.evaluation_dir,a.tessdata_dir,[x.strip() for x in a.models.split(",") if x.strip()],a.tesseract,a.psm)
    except (OSError,ValueError,RuntimeError,subprocess.TimeoutExpired) as e:print(f"ERROR: {e}",file=sys.stderr);return 2
    if not report["ranking"]:print("ERROR: no requested model could be evaluated",file=sys.stderr);return 2
    a.output_dir.mkdir(parents=True,exist_ok=True);csv_path=a.output_dir/"ocr-model-comparison.csv"
    with csv_path.open("w",encoding="utf-8-sig",newline="") as f:
        writer=csv.DictWriter(f,fieldnames=["model","image","expected","predicted","cer","wer"]);writer.writeheader();writer.writerows(rows)
    json_path=a.output_dir/"ocr-model-summary.json";json_path.write_text(json.dumps(report,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    print(json.dumps(report,ensure_ascii=False,indent=2));print(f"CSV: {csv_path}\nJSON: {json_path}");return 0
if __name__=="__main__":raise SystemExit(main())
