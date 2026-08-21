# -*- coding: utf-8 -*-
"""Y-walk: match Word PDF lines against an Oxi --dump-layout by text and print
the running dy, so a page-count difference caused by ACCUMULATED drift can be
localized to the row where it enters.

  python tools/metrics/_ywalk.py <docx> [--reexport]

Word y is the PDF glyph top; Oxi y is the line-box top plus text_y_off, so the
absolute offset is a constant convention gap — read the SLOPE, not the value.
"""
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
DOCX = Path(sys.argv[1]).resolve()
PDF = Path(tempfile.gettempdir()) / (DOCX.name + ".truth.pdf")


def word_lines():
    if not PDF.exists() or "--reexport" in sys.argv:
        import win32com.client as win32
        w = win32.DispatchEx("Word.Application")
        w.Visible = False
        try:
            d = w.Documents.Open(str(DOCX), ReadOnly=True)
            d.ExportAsFixedFormat(str(PDF), 17)
            d.Close(False)
        finally:
            w.Quit()
    import fitz
    doc = fitz.open(PDF)
    out = []
    for pi in range(doc.page_count):
        for blk in doc[pi].get_text("dict")["blocks"]:
            if blk.get("type", 0) != 0:
                continue
            for ln in blk.get("lines", []):
                t = "".join(s["text"] for s in ln["spans"])
                if not t.strip():
                    continue
                out.append((pi, min(s["bbox"][1] for s in ln["spans"]),
                            min(s["bbox"][0] for s in ln["spans"]), t.strip()))
    out.sort()
    return out


def oxi_lines():
    exe = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
    tmp = Path(tempfile.mkdtemp())
    dump = tmp / "d.json"
    subprocess.run([str(exe), str(DOCX), str(tmp / "p"), "110", "--dump-layout=%s" % dump],
                   check=True, capture_output=True)
    import json
    d = json.load(open(dump, encoding="utf-8"))
    out = []
    for pi, pg in enumerate(d["pages"]):
        rows = {}
        for e in pg["elements"]:
            if e.get("type") != "text":
                continue
            t = e.get("text") or ""
            if not t.strip():
                continue
            key = (round(e.get("y", 0.0), 1), round(e.get("x", 0.0) / 40))
            rows.setdefault(key, []).append((e.get("x", 0.0), t))
        for (y, _), frags in rows.items():
            frags.sort()
            out.append((pi, y, frags[0][0], "".join(f[1] for f in frags).strip()))
    out.sort()
    return out


def norm(s):
    return re.sub(r"[^0-9a-z]", "", s.lower())[:20]


w = word_lines()
o = oxi_lines()
print("word lines %d (pages %d) / oxi lines %d (pages %d)"
      % (len(w), max(r[0] for r in w) + 1, len(o), max(r[0] for r in o) + 1))
used = set()
print("  %-4s %-9s %-9s %-8s %s" % ("pg", "word_y", "oxi_y", "dy", "text"))
base = None
for pi, wy, wx, wt in w:
    k = norm(wt)
    if len(k) < 6:
        continue
    cand = [(j, r) for j, r in enumerate(o) if j not in used and r[0] == pi and norm(r[3]) == k]
    if not cand:
        continue
    j, r = cand[0]
    used.add(j)
    dy = r[1] - wy
    if base is None:
        base = dy
    print("  %-4d %-9.2f %-9.2f %+8.2f  %s" % (pi, wy, r[1], dy - base, wt[:44]))
