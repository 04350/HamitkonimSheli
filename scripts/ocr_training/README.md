# אימון מודל OCR מותאם למתכונים

התשתית מאמנת שתי גרסאות על אותו מאגר מתכונים: אחת מתחילה מ־`heb` ואחת מ־`heb4`. לאחר מכן משווים אותן מול מודלי הבסיס על סריקות אחרות, שלא שימשו לאימון.

**התמונות והתמלולים נשארים מקומית. אין להעלות אותם ל־GitHub.** תיקיות הנתונים הפרטיות מוחרגות מ־Git.

## 1. סביבת עבודה

נדרש Linux או WSL2 עם Tesseract 5 וכלי LSTM. ב־Ubuntu:

```bash
sudo apt update
sudo apt install -y git make wget curl python3 python3-venv tesseract-ocr
python3 -m venv .venv-ocr
source .venv-ocr/bin/activate
python -m pip install --upgrade pip pillow pymupdf
```

אחרי פתיחת סביבת העבודה בכל טרמינל חדש, הפעל קודם: `source .venv-ocr/bin/activate`. בדוק שהפקודות `tesseract`, `lstmtraining`, `combine_tessdata`, `make` ו־`wget` זמינות.

## 2. סימון שורות

מתוך שורש הפרויקט:

```bash
python3 scripts/ocr_training/annotate.py --split training
```

פתח סריקה או PDF, גרור עם העכבר מלבן סביב שורת טקסט אחת, הקלד את הטקסט בדיוק כפי שמופיע ולחץ "שמור שורה". נוצרים שני קבצים תואמים:

- `recipe-line-00001.png` — תמונת השורה.
- `recipe-line-00001.gt.txt` — התמלול המדויק, UTF-8 ובשורה אחת.

דוגמאות בדיקה מסמנים בנפרד:

```bash
python3 scripts/ocr_training/annotate.py --split evaluation
```

חשוב: תיקיית הבדיקה צריכה להכיל עמודים, ספרים או צילומים שלא הופיעו באימון. אין לפצל שורות מאותו עמוד בין אימון לבדיקה.

כלול מגוון גופנים ואיכויות סריקה, עברית ואנגלית, כותרות, שמות רכיבים, מספרים, שברים כמו 1/2, מידות כמו גרם ומ"ל, סימני פיסוק וטקסט משולב. יש להקליד רק את הטקסט המדויק שנראה בסריקה.

## 3. אימות הנתונים

```bash
python3 scripts/ocr_training/validate_dataset.py data/recipe-ocr-training --min-samples 20
python3 scripts/ocr_training/validate_dataset.py data/recipe-ocr-evaluation --min-samples 10
```

## 4. אימון והשוואה

```bash
bash scripts/ocr_training/run_training.sh
```

אפשר לשנות את מספר האיטרציות:

```bash
MAX_ITERATIONS=15000 bash scripts/ocr_training/run_training.sh
```

הסקריפט מוריד את tesstrain ואת קובצי המודל הדרושים בשלב הפיתוח, מאמן את `recipe_heb` ואת `recipe_heb4`, ומשווה אותם למודלי הבסיס על מאגר הבדיקה. המדדים כוללים CER (שגיאות תווים), WER (שגיאות מילים), שגיאות ספרות ודיוק שורה מלאה. התוצאות נשמרות תחת `build/recipe-ocr-training`.

מינימום 20/10 שורות נועד רק לוודא שהצינור פועל; הוא אינו מספיק להכרעה שמודל טוב יותר. להחלטת הטמעה כדאי לפחות 100 שורות בדיקה ממקורות מגוונים ו־500–1,000 שורות אימון ומעלה.

## 5. בחירת מודל והטמעה

`ocr-model-summary.json` מציג דירוג ו־`ocr-model-comparison.csv` מפרט כל זיהוי. CER הוא שיעור שגיאות התווים, WER הוא שיעור שגיאות המילים ומדד הספרות בודק מספרים בנפרד.

אין החלפה אוטומטית של מודל התוכנה. בוחרים מועמד רק אם השיפור עקבי ואינו פוגע בזיהוי כמויות, שברים ומידות. לאחר בחירה אפשר להעתיק את `recipe_heb.traineddata` או `recipe_heb4.traineddata` אל `tesseract/tessdata`, להתאים את קוד השפה ולבנות חבילת Windows חדשה. עד אז מנוע הייצור נשאר ללא שינוי.

## פרטי מימוש ורישוי

- נעשה שימוש ב־tesstrain הרשמי וב־Fine-tuning של Tesseract LSTM; לא מחברים בינארית קובצי מודל.
- `heb4` מגיע מ־`boro11204/hebrew-ocr-offline`, רוויזיה `30ae3e5513d5c132ecdfeb03730fa0b723a2180f`; הסקריפט מאמת SHA-256.
- מדדי המקור של `heb4` הם על חומר מודפס מסוג אחר, ואינם מוכיחים דיוק במתכונים.
- האימון הוא שלב פיתוח. בזמן שימוש בתוכנת Windows הזיהוי נשאר מקומי ואינו דורש רשת.
