#!/bin/sh
# Annotate docling's own JSON for every .docx in a folder, then check the result.
#   scripts/export_all.sh <docx dir> <docling json dir> <out dir> <DoclingDocument_1_10.json>
# A docling JSON is looked up as <name>.docx.json (docling's ground-truth naming), then <name>.json.
set -e
HERE="$(dirname "$0")"
BIN="$HERE/../../../target/release/oxi-docling"
mkdir -p "$3"
for f in "$1"/*.docx; do
  n="$(basename "$f" .docx)"
  j="$2/$n.docx.json"; [ -f "$j" ] || j="$2/$n.json"
  [ -f "$j" ] && "$BIN" annotate "$f" "$j" "$3/$n.json" 2>>"$3/_stderr.txt"
done
python3 "$HERE/validate_docling.py" "$4" "$3"/*.json
python3 "$HERE/check_annotate.py" "$2" "$3"
