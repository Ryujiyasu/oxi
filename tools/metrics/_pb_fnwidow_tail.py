# -*- coding: utf-8 -*-
"""Print the tail of page 1 (and head of page 2) for one arm, Word vs Oxi.

Usage: python tools/metrics/_pb_fnwidow_tail.py f46_r3_p2
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fnwidow_gen import OUT

REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
tag = sys.argv[1]
docx = os.path.join(OUT, tag + ".docx")


def word_pages(docx):
    import fitz
    doc = fitz.open(docx[:-5] + ".pdf")
    out = []
    for pno in range(min(2, doc.page_count)):
        pg = doc[pno]
        sep = None
        for dr in pg.get_drawings():
            r = dr["rect"]
            if r.height < 4 and 100 < r.width < 200 and r.y0 > 300:
                sep = r.y0 if sep is None else min(sep, r.y0)
        rows = []
        for blk in pg.get_text("dict")["blocks"]:
            for l in blk.get("lines", []):
                sp = [s for s in l["spans"] if s["text"].strip()]
                if sp:
                    rows.append((round(sp[0]["origin"][1], 2),
                                 "".join(s["text"] for s in sp).strip()))
        rows.sort()
        out.append((rows, sep))
    return out


def oxi_pages(docx):
    dump = docx[:-5] + ".layout.json"
    if not os.path.exists(dump):
        subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                       capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    out = []
    for pg in d["pages"][:2]:
        sep = None
        for e in pg["elements"]:
            if (e.get("text") or "").strip():
                continue
            if e.get("y", 0) > 300 and 100 < e.get("w", 0) < 200 and e.get("h", 9) < 4:
                sep = e["y"] if sep is None else min(sep, e["y"])
        rows = {}
        for e in pg["elements"]:
            if not (e.get("text") or "").strip():
                continue
            rows.setdefault(round(e["y"] + 10.5, 2), []).append((e.get("x", 0), e["text"]))
        out.append(([(y, "".join(t for _, t in sorted(rows[y])).strip())
                     for y in sorted(rows)], sep))
    return out


for name, pages in (("WORD", word_pages(docx)), ("OXI", oxi_pages(docx))):
    print("=== %s  %s ===" % (name, tag))
    for pno, (rows, sep) in enumerate(pages):
        print("  page %d   sep=%s" % (pno + 1, sep))
        for y, t in rows:
            mark = "  note" if (sep is not None and y > sep) else "      "
            print("    %8.2f%s  %s" % (y, mark, t[:58]))
        print()
