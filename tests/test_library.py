# -*- coding: utf-8 -*-
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import ספריית_המתכונים as app

class בדיקותספרייה(unittest.TestCase):
    def test_ניקוי_חיפוש_עברית_וניקוד(self):
        self.assertEqual(app.normalize("עוגת שׁוֹקוֹלָד!"), app.normalize("עוגת שוקולד"))

    def test_שם_קובץ_בטוח(self):
        self.assertEqual(app.safe_title('עוגה: "טעימה"?'), "עוגה טעימה")

    def test_זיהוי_כותרת_מתוך_טקסט(self):
        self.assertEqual(app.choose_title(Path("קובץ.txt"), "עוגת תפוחים\nמערבבים את החומרים"), "עוגת תפוחים")

    def test_חיפוש_נותן_עדיפות_לשם(self):
        name_score = app.relevance("עוגת שוקולד", "מתכון פשוט", "עוגת שוקולד")
        body_score = app.relevance("מתכון אחר", "עוגת שוקולד טעימה", "עוגת שוקולד")
        self.assertGreater(name_score, body_score)

    def test_שם_ללא_תוכן_נשאר_קריא(self):
        self.assertGreater(app.relevance("קובץ תמונה", "", "תמונה"), 0)

if __name__ == "__main__":
    unittest.main()
