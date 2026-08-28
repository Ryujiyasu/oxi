# -*- coding: utf-8 -*-
"""Read the widow/orphan probe. Usage: _pb_widow_read.py word|oxi"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_widow_gen import ARMS, OUT, SHAPES, FILLS, TALL
from _pb_kntbl_read import word_pages
import json, subprocess
REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")

def oxi_pages(docx):
    """Same as _pb_kntbl_read.oxi_pages but joins same-baseline runs with a
    SPACE.  Oxi emits a justified body line one word per run, and joining those
    with "" welds them into a single token no word matcher can read."""
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8")); out = []
    for pi, pg in enumerate(d["pages"], 1):
        rows = {}
        for e in pg["elements"]:
            if not (e.get("text") or "").strip(): continue
            rows.setdefault(round(e["y"], 1), []).append((round(e["x"], 1), e["text"]))
        for y, frs in rows.items():
            frs.sort(); out.append((pi, y, " ".join(t for _, t in frs)))
    return sorted(out)
WORDS = TALL.split()

def split_shape(lines):
    """Lines of the probe paragraph per page.

    Both readers concatenate everything sharing a y: the cell arms pick up the
    neighbour cell's "B", and Oxi emits the body arms one word per run.  Keep
    only words belonging to the probe string and require a contiguous slice, so
    a line is counted wherever it is drawn and whatever shares its baseline.
    """
    per = {}
    for p, y, t in lines:
        w = [x for x in t.replace("|", " ").split() if x in WORDS]
        if not w or len(w) > len(WORDS): continue
        n = len(w)
        if any(WORDS[i:i + n] == w for i in range(len(WORDS) - n + 1)):
            per[p] = per.get(p, 0) + 1
    return per

mode = sys.argv[1] if len(sys.argv) > 1 else "word"
res = {}
for tag, n, blk, off in ARMS:
    docx = os.path.join(OUT, tag + ".docx")
    res[tag] = split_shape(word_pages(docx) if mode == "word" else oxi_pages(docx))
print(f"{mode.upper()}  lines of the 5-line paragraph left on p1 / carried to p2")
print("  fill  " + "".join(f"{s:>12s}" for s in SHAPES))
for n in FILLS:
    cells = []
    for s in SHAPES:
        per = res[f"{s}{n}"]
        cells.append("/".join(str(per.get(p, 0)) for p in (1, 2)))
    print(f"  {n:4d}  " + "".join(f"{c:>12s}" for c in cells))
