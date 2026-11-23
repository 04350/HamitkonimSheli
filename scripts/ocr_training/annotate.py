#!/usr/bin/env python3
"""Mark recipe text lines and save Tesseract PNG/.gt.txt training pairs."""
from __future__ import annotations
import argparse, re, tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageOps, ImageTk
try:
    import fitz
except Exception:
    fitz=None
IMAGE_SUFFIXES={".png",".jpg",".jpeg",".tif",".tiff",".bmp",".webp"}

class Annotator:
    def __init__(self, output: Path):
        self.output=output; self.output.mkdir(parents=True,exist_ok=True)
        self.pages=[]; self.index=0; self.scale=1.0; self.ox=self.oy=0
        self.start=None; self.rect=None; self.selection=None; self.preview=None
        self.number=self.next_number()
        self.root=tk.Tk(); self.root.title("סימון שורות לאימון OCR של מתכונים"); self.root.geometry("1200x860")
        bar=ttk.Frame(self.root,padding=8); bar.pack(fill="x")
        ttk.Button(bar,text="פתיחת תמונות / PDF",command=self.open_files).pack(side="right",padx=3)
        ttk.Button(bar,text="הקודם",command=self.prev).pack(side="right",padx=3)
        ttk.Button(bar,text="הבא",command=self.next).pack(side="right",padx=3)
        ttk.Button(bar,text="תיקיית שמירה",command=self.choose_output).pack(side="right",padx=3)
        self.info=ttk.Label(bar,text="פתח סריקות מתכונים כדי להתחיל"); self.info.pack(side="left",padx=8)
        self.canvas=tk.Canvas(self.root,bg="#242424",highlightthickness=0); self.canvas.pack(fill="both",expand=True,padx=8,pady=5)
        self.canvas.bind("<Configure>",lambda _e:self.show())
        self.canvas.bind("<ButtonPress-1>",self.begin); self.canvas.bind("<B1-Motion>",self.drag); self.canvas.bind("<ButtonRelease-1>",self.end)
        bottom=ttk.Frame(self.root,padding=8); bottom.pack(fill="x")
        ttk.Label(bottom,text="הקלד תמלול מדויק של השורה המסומנת:").pack(anchor="e")
        self.text=ttk.Entry(bottom,font=("Segoe UI",13),justify="right"); self.text.pack(fill="x",pady=6)
        self.text.bind("<Return>",lambda _e:self.save())
        ttk.Button(bottom,text="שמור שורה",command=self.save).pack(side="right",padx=4)
        ttk.Button(bottom,text="נקה סימון",command=self.clear).pack(side="right",padx=4)
        self.status=ttk.Label(bottom,text=f"שמירה אל: {self.output}"); self.status.pack(side="left",padx=8)

    def next_number(self):
        ids=[int(m.group(1)) for p in self.output.glob("recipe-line-*.png") if (m:=re.search(r"(\d+)$",p.stem))]
        return max(ids,default=0)+1

    def open_files(self):
        names=filedialog.askopenfilenames(title="בחירת סריקות מתכונים",filetypes=[
            ("תמונות ו-PDF","*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp *.pdf"),("כל הקבצים","*.*")])
        if not names:return
        pages=[]; errors=[]
        for name in names:
            path=Path(name)
            try:
                if path.suffix.lower()==".pdf":
                    if fitz is None: raise RuntimeError("להצגת PDF נדרש PyMuPDF: pip install pymupdf")
                    doc=fitz.open(path)
                    try:
                        for n,page in enumerate(doc,1):
                            pix=page.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False)
                            pages.append((str(path),n,Image.frombytes("RGB",[pix.width,pix.height],pix.samples)))
                    finally: doc.close()
                elif path.suffix.lower() in IMAGE_SUFFIXES:
                    with Image.open(path) as im: pages.append((str(path),1,ImageOps.exif_transpose(im).convert("RGB")))
                else: errors.append(f"סוג קובץ לא נתמך: {path.name}")
            except Exception as exc: errors.append(f"{path.name}: {exc}")
        for _src,_page,im in self.pages: im.close()
        self.pages=pages; self.index=0; self.show()
        if errors: messagebox.showwarning("קבצים שלא נפתחו","\n".join(errors[:8]))

    def show(self):
        if not self.pages:return
        source,page_no,im=self.pages[self.index]
        w,h=max(100,self.canvas.winfo_width()-20),max(100,self.canvas.winfo_height()-20)
        self.scale=min(w/im.width,h/im.height,1.0)
        rw,rh=max(1,int(im.width*self.scale)),max(1,int(im.height*self.scale))
        thumb=im.resize((rw,rh),Image.Resampling.LANCZOS); self.preview=ImageTk.PhotoImage(thumb)
        self.ox=(self.canvas.winfo_width()-rw)//2; self.oy=(self.canvas.winfo_height()-rh)//2
        self.canvas.delete("all"); self.canvas.create_image(self.ox,self.oy,anchor="nw",image=self.preview)
        self.rect=None; self.selection=None
        self.info.configure(text=f"{Path(source).name} · עמוד {page_no} · {self.index+1}/{len(self.pages)}")

    def coords(self,x,y):
        im=self.pages[self.index][2]
        return max(0,min(im.width,int((x-self.ox)/self.scale))),max(0,min(im.height,int((y-self.oy)/self.scale)))

    def begin(self,event):
        if not self.pages:return
        self.clear(); self.start=self.coords(event.x,event.y)

    def drag(self,event):
        if self.start is None:return
        x1=self.ox+self.start[0]*self.scale; y1=self.oy+self.start[1]*self.scale
        if self.rect is not None:self.canvas.delete(self.rect)
        self.rect=self.canvas.create_rectangle(x1,y1,event.x,event.y,outline="#00d084",width=2)

    def end(self,event):
        if self.start is None:return
        x2,y2=self.coords(event.x,event.y); x1,y1=self.start
        self.selection=(min(x1,x2),min(y1,y2),max(x1,x2),max(y1,y2))
        self.start=None; self.text.focus_set()

    def clear(self):
        if self.rect is not None:self.canvas.delete(self.rect)
        self.rect=None; self.selection=None; self.start=None

    def save(self):
        if not self.pages or self.selection is None:
            messagebox.showinfo("חסר סימון","סמן שורת טקסט תחילה."); return
        text=self.text.get().strip()
        if not text or "\n" in text or "\r" in text:
            messagebox.showinfo("תמלול לא תקין","יש להזין שורה אחת מלאה ומדויקת."); return
        x1,y1,x2,y2=self.selection
        if x2-x1<10 or y2-y1<8:
            messagebox.showinfo("בחירה קטנה מדי","סמן את כל השורה בלי לחתוך אותיות."); return
        crop=self.pages[self.index][2].crop((x1,y1,x2,y2)).convert("L")
        crop=ImageOps.expand(crop,border=(12,8),fill=255)
        stem=f"recipe-line-{self.number:05d}"; crop.save(self.output/f"{stem}.png",optimize=True); crop.close()
        (self.output/f"{stem}.gt.txt").write_text(text+"\n",encoding="utf-8")
        self.number+=1; self.text.delete(0,"end"); self.clear()
        self.status.configure(text=f"נשמרו {self.number-1} שורות · {self.output}")

    def prev(self):
        if self.index>0:self.index-=1; self.show()
    def next(self):
        if self.index+1<len(self.pages):self.index+=1; self.show()
    def choose_output(self):
        name=filedialog.askdirectory(title="בחירת תיקיית שמירה")
        if name:
            self.output=Path(name); self.output.mkdir(parents=True,exist_ok=True)
            self.number=self.next_number(); self.status.configure(text=f"שמירה אל: {self.output}")
    def run(self):
        self.root.mainloop()
        for _src,_page,im in self.pages: im.close()

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--split",choices=("training","evaluation"),default="training")
    parser.add_argument("--output-dir",type=Path)
    args=parser.parse_args()
    root=Path(__file__).resolve().parents[2]
    default=root/"data"/("recipe-ocr-training" if args.split=="training" else "recipe-ocr-evaluation")
    Annotator(args.output_dir or default).run()
if __name__=="__main__": main()
