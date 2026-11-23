#!/usr/bin/env bash
# Fine-tune both Hebrew starting models with identical recipe data and compare.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "$HERE/../.." && pwd)"
TRAIN_DIR="${TRAIN_DIR:-$ROOT/data/recipe-ocr-training}"
EVAL_DIR="${EVAL_DIR:-$ROOT/data/recipe-ocr-evaluation}"
WORK="${WORK_DIR:-$ROOT/build/recipe-ocr-training}"
ITERATIONS="${MAX_ITERATIONS:-10000}"
REF="${TESSTRAIN_REF:-405346a3a67d8e4e049341d1da6a4b752e0b8351}"
TESSDATA="$WORK/tessdata"
TESSTRAIN="$WORK/tesstrain"
for cmd in python3 make git curl sha256sum tesseract lstmtraining combine_tessdata wget; do
  command -v "$cmd" >/dev/null 2>&1 || { echo "Missing command: $cmd" >&2; exit 2; }
done
python3 "$HERE/validate_dataset.py" "$TRAIN_DIR" --min-samples 20
python3 "$HERE/validate_dataset.py" "$EVAL_DIR" --min-samples 10
mkdir -p "$TESSDATA" "$WORK"
if [[ ! -d "$TESSTRAIN/.git" ]]; then git clone https://github.com/tesseract-ocr/tesstrain.git "$TESSTRAIN"; fi
git -C "$TESSTRAIN" fetch --depth 1 origin "$REF"
git -C "$TESSTRAIN" checkout --detach FETCH_HEAD
curl -fL --retry 3 https://raw.githubusercontent.com/tesseract-ocr/tessdata_best/e12c65a915945e4c28e237a9b52bc4a8f39a0cec/heb.traineddata -o "$TESSDATA/heb.traineddata"
curl -fL --retry 3 https://raw.githubusercontent.com/boro11204/hebrew-ocr-offline/30ae3e5513d5c132ecdfeb03730fa0b723a2180f/models/heb4.traineddata -o "$TESSDATA/heb4.traineddata"
printf '%s  %s\n' 'a1422a71e9e7d060b5cc0c08e3060cc5e7a9a9ea41ecc0da43ec07d8946769a6' "$TESSDATA/heb4.traineddata" | sha256sum --check --status || { echo "heb4 checksum verification failed" >&2; exit 3; }
curl -fL --retry 3 https://raw.githubusercontent.com/tesseract-ocr/tessdata_best/e12c65a915945e4c28e237a9b52bc4a8f39a0cec/eng.traineddata -o "$TESSDATA/eng.traineddata"
export TESSDATA_PREFIX="$TESSDATA"
train_candidate() {
  local base="$1" name="$2" data="$WORK/data" output="$WORK/data/$2"
  echo "Fine-tuning $name from $base ..."
  # tesstrain requires OUTPUT_DIR to be exactly DATA_DIR/MODEL_NAME.
  make -C "$TESSTRAIN" training MODEL_NAME="$name" START_MODEL="$base" TESSDATA="$TESSDATA" \
    DATA_DIR="$data" LANGDATA_DIR="$WORK/langdata" \
    GROUND_TRUTH_DIR="$TRAIN_DIR" MAX_ITERATIONS="$ITERATIONS" LANG_TYPE=RTL PSM=7 RATIO_TRAIN=0.90
  test -s "$output.traineddata"
  cp "$output.traineddata" "$TESSDATA/$name.traineddata"
}
train_candidate heb recipe_heb
train_candidate heb4 recipe_heb4
python3 "$HERE/evaluate_models.py" --evaluation-dir "$EVAL_DIR" --tessdata-dir "$TESSDATA" \
  --models 'heb+eng,heb4+eng,recipe_heb+eng,recipe_heb4+eng' --output-dir "$WORK/evaluation"
echo "Candidates: $TESSDATA/recipe_heb.traineddata, $TESSDATA/recipe_heb4.traineddata"
echo "See $WORK/evaluation/ocr-model-summary.json and ocr-model-comparison.csv"
echo "Do not deploy a candidate until held-out recipe scans prove a repeatable gain."
