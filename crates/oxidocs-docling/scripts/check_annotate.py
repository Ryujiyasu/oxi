"""Check annotate-mode output against the docling JSON it was made from.

    python check_annotate.py <docling json dir> <annotated dir>

For every pair (same file name; a docling ground-truth name `x.docx.json`
pairs with `x.json`): every text, table and picture must carry `prov`, and
nothing but `prov`, `pages` and table cell `bbox` may differ from the input.
"""
import copy
import json
import sys
from pathlib import Path


def strip(d):
    d = copy.deepcopy(d)
    d["pages"] = {}
    for kind in ("texts", "tables", "pictures"):
        for it in d[kind]:
            it["prov"] = []
            if kind == "tables":
                for c in it["data"]["table_cells"]:
                    c.pop("bbox", None)
    return d


src_dir, out_dir = Path(sys.argv[1]), Path(sys.argv[2])
files = items = placed = changed = 0
for out in sorted(out_dir.glob("*.json")):
    src = src_dir / (out.stem + ".docx.json")
    if not src.exists():
        src = src_dir / out.name
    if not src.exists():
        continue
    a, d = json.loads(out.read_text()), json.loads(src.read_text())
    files += 1
    for kind in ("texts", "tables", "pictures"):
        for it in a[kind]:
            items += 1
            placed += bool(it["prov"])
    if strip(a) != strip(d):
        changed += 1
        print(f"CHANGED beyond provenance: {out.name}")
print(f"{files} files, {placed}/{items} items with prov, {changed} files changed beyond provenance")
