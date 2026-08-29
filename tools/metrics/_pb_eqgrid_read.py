# -*- coding: utf-8 -*-
"""Read the display-equation grid probe. Usage: _pb_eqgrid_read.py word|oxi

Prints, per arm, the advance from the last BEFORE line to the first AFTER line
and that advance in grid cells. Ceil-to-grid shows as a whole number in every
arm; a constant shows as the same POINT value regardless of the equation.
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_eqgrid_gen import ARMS, OUT, GRIDS, SHAPES

REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")


def word_rows(docx):
    import fitz, win32com.client
    pdf = docx[:-5] + ".pdf"
    if not os.path.exists(pdf):
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            d.SaveAs2(os.path.abspath(pdf), FileFormat=17)
            d.Close(False)
        finally:
            word.Quit()
    doc = fitz.open(pdf)
    before, after = [], []
    for pno in range(doc.page_count):
        for blk in doc[pno].get_text("dict")["blocks"]:
            for l in blk.get("lines", []):
                t = "".join(s["text"] for s in l["spans"]).strip()
                if t.startswith("BEFORE"):
                    before.append(l["bbox"][1])
                elif t.startswith("AFTER"):
                    after.append(l["bbox"][1])
    return (max(before) if before else None), (min(after) if after else None)


def oxi_rows(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    before, after = [], []
    for pg in d["pages"]:
        rows = {}
        for e in pg["elements"]:
            t = (e.get("text") or "")
            if not t.strip():
                continue
            rows.setdefault(round(e["y"], 2), []).append((e.get("x", 0), t))
        for y, items in rows.items():
            s = "".join(t for _, t in sorted(items))
            if s.startswith("BEFORE"):
                before.append(y)
            elif s.startswith("AFTER"):
                after.append(y)
    return (max(before) if before else None), (min(after) if after else None)


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_rows if mode == "word" else oxi_rows
print("%s   advance across the display equation\n" % mode.upper())
print("  grid   shape    advance   cells")
for tag, pitch, _ in ARMS:
    docx = os.path.join(OUT, tag + ".docx")
    if not os.path.exists(docx):
        print("  %-18s MISSING" % tag)
        continue
    b, a = reader(docx)
    cell = pitch / 20.0
    if b is None or a is None:
        print("  %-18s  (no BEFORE/AFTER found)" % tag)
        continue
    adv = a - b
    print("  %-6s %-8s %7.2f   %6.3f" % (tag.split("_")[0], tag.split("_")[1], adv, adv / cell))
