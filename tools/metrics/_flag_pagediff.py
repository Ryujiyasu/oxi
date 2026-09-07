# -*- coding: utf-8 -*-
"""Render one document twice with the GDI renderer -- once with an environment
flag set, once without -- and print the lines that differ on the given pages,
plus the Word PDF's lines there when a PDF is found next to the docx or under
pipeline_data. Finds which paragraph a flag pushes onto another line.

Usage: python tools/metrics/_flag_pagediff.py <docx> FLAG[=value] page [page ...]
"""
import difflib
import glob
import json
import os
import subprocess
import sys
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
EXE = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")
docx = os.path.abspath(sys.argv[1])
flag, _, value = sys.argv[2].partition("=")
pages = [int(p) for p in sys.argv[3:]]
TMP = os.path.join(os.environ.get("TMP", "."), "flag_pagediff")
os.makedirs(TMP, exist_ok=True)


def dump(env, name):
    out = os.path.join(TMP, name + ".json")
    e = dict(os.environ)
    e.update(env)
    subprocess.run([EXE, docx, os.path.join(TMP, name), "--dump-layout=" + out], env=e, capture_output=True)
    return json.load(open(out, encoding="utf-8"))["pages"]


def lines_of(page):
    byy = defaultdict(list)
    for el in page.get("elements") or []:
        t = el.get("text") or ""
        if t.strip():
            byy[round(el.get("y", -1), 1)].append((el.get("x", -1), t))
    out = []
    for y in sorted(byy):
        items = sorted(byy[y])
        out.append((y, items[0][0], "".join(t for _, t in items)))
    return out


on = dump({flag: value or "1"}, "on")
off = dump({}, "off")
for pno in pages:
    a = lines_of(off[pno - 1]) if pno - 1 < len(off) else []
    b = lines_of(on[pno - 1]) if pno - 1 < len(on) else []
    print("== Oxi p%d: default %d lines / %s %d lines" % (pno, len(a), flag, len(b)))
    ta = [t for _, _, t in a]
    tb = [t for _, _, t in b]
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(a=ta, b=tb, autojunk=False).get_opcodes():
        if tag == "equal":
            # same text: report a horizontal shift
            for k in range(i2 - i1):
                (ya, xa, t), (yb, xb, _) = a[i1 + k], b[j1 + k]
                if abs(xa - xb) > 0.05 or abs(ya - yb) > 0.05:
                    print("   SHIFT y=%6.1f x %.2f -> %.2f (%+.2f) dy=%+.2f n=%2d %s" % (ya, xa, xb, xb - xa, yb - ya, len(t), t[:40]))
            continue
        for y, x, t in a[i1:i2]:
            print("   OFF y=%6.1f x=%6.1f n=%2d %s" % (y, x, len(t), t[:46]))
        for y, x, t in b[j1:j2]:
            print("   ON  y=%6.1f x=%6.1f n=%2d %s" % (y, x, len(t), t[:46]))
stem = os.path.splitext(os.path.basename(docx))[0]
cands = glob.glob(os.path.join(os.path.dirname(docx), stem + "*.pdf")) + glob.glob(os.path.join(REPO, "pipeline_data", "**", stem[:12] + "*.pdf"), recursive=True)
cands = [c for c in cands if "lo_pdf" not in c and "libra" not in c and "onlyoffice" not in c]
if cands:
    import fitz
    d = fitz.open(cands[0])
    print("== Word PDF:", cands[0])
    for pno in pages:
        if pno - 1 >= len(d):
            continue
        rows = []
        for bl in d[pno - 1].get_text("dict")["blocks"]:
            for l in bl.get("lines", []):
                t = "".join(s["text"] for s in l["spans"])
                if t.strip():
                    rows.append((round(l["bbox"][1], 1), round(l["bbox"][0], 1), t))
        rows.sort()
        print("   p%d: %d lines" % (pno, len(rows)))
        for y, x, t in rows:
            print("     y=%6.1f x=%6.1f n=%2d %s" % (y, x, len(t.strip()), t[:46]))
