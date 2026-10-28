#!/bin/sh
# Export every .docx in a folder and run the two checks.
#   scripts/export_all.sh <docx dir> <out dir> <DoclingDocument_1_10.json> [groundtruth dir]
set -e
BIN="$(dirname "$0")/../../../target/release/oxi-docling"
mkdir -p "$2"
for f in "$1"/*.docx; do "$BIN" "$f" "$2/$(basename "$f" .docx).json" 2>>"$2/_stderr.txt"; done
python3 "$(dirname "$0")/validate_docling.py" "$3" "$2"/*.json
[ -n "$4" ] && python3 "$(dirname "$0")/compare_docling.py" "$2" "$4"
