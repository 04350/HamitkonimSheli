from __future__ import annotations
import importlib.util, tempfile, unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
SCRIPTS=ROOT/"scripts"/"ocr_training"
def load(name,path):
    spec=importlib.util.spec_from_file_location(name,path)
    module=importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module
validate=load("validate_dataset",SCRIPTS/"validate_dataset.py")
evaluate=load("evaluate_models",SCRIPTS/"evaluate_models.py")
class OCRTrainingTests(unittest.TestCase):
    def test_matching_training_pair_is_valid(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);(root/"line-1.png").write_bytes(b"image")
            (root/"line-1.gt.txt").write_text("2 כוסות קמח\n",encoding="utf-8")
            report=validate.inspect_dataset(root)
            self.assertEqual(report["valid_samples"],1)
            self.assertEqual(report["missing_transcriptions"],[])
    def test_missing_pair_and_multiline_are_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);(root/"line-1.png").write_bytes(b"image")
            (root/"line-2.png").write_bytes(b"image")
            (root/"line-2.gt.txt").write_text("שורה א\nשורה ב",encoding="utf-8")
            report=validate.inspect_dataset(root)
            self.assertEqual(report["missing_transcriptions"],["line-1"])
            self.assertEqual(report["multiline_transcriptions"],["line-2.gt.txt"])
    def test_levenshtein_counts_character_and_digit_errors(self):
        self.assertEqual(evaluate.distance("עוגה","עוגה"),0)
        self.assertEqual(evaluate.distance("1/2","1/3"),1)
        self.assertEqual(evaluate.distance("123","12"),1)
    def test_normalization_preserves_numbers(self):
        self.assertEqual(evaluate.norm("2  כוסות\n"),"2 כוסות")
        self.assertNotEqual(evaluate.norm("1/2"),evaluate.norm("1/3"))
if __name__=="__main__":unittest.main()
