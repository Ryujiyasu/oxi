# -*- coding: utf-8 -*-
"""Read the hanging-punctuation probe. Usage: _pb_hangpunct_read.py word|oxi

Per arm: how many lines the filler paragraph took, and how far the last glyph
reaches. A mark that may hang keeps the paragraph on ONE line at the filler
length that exactly fills the column, and its glyph ends past the right edge
(523.32pt).
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_hangpunct_gen import ARMS, OUT, PUNCT, NS, JC

RIGHT = 72.0 + 451.32
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
    lines = []
    for blk in doc[0].get_text("rawdict")["blocks"]:
        for l in blk.get("lines", []):
            chs = [c for s in l["spans"] for c in s.get("chars", []) if c["c"].strip()]
            if not chs:
                continue
            txt = "".join(c["c"] for c in chs)
            if txt.startswith("AFTER"):
                continue
            lines.append((round(l["bbox"][1], 2), max(c["bbox"][2] for c in chs)))
    lines.sort()
    return lines


def oxi_rows(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    rows = {}
    for pg in d["pages"]:
        for e in pg["elements"]:
            t = (e.get("text") or "")
            if not t.strip() or t.strip().startswith("AFTER"):
                continue
            y = round(e["y"], 2)
            rows.setdefault(y, []).append(e.get("x", 0) + e.get("w", 0))
    return sorted((y, max(v)) for y, v in rows.items())


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_rows if mode == "word" else oxi_rows
print("%s   filler + one mark, column right edge = %.2fpt\n" % (mode.upper(), RIGHT))
print("  %-8s %-8s %s" % ("jc", "mark", "  ".join("n=%d lines/right" % n for n in NS)))
for j in JC:
    for p in PUNCT:
        cells = []
        for n in NS:
            docx = os.path.join(OUT, "%s_%s_n%d.docx" % (j, p, n))
            if not os.path.exists(docx):
                cells.append("   --      ")
                continue
            r = reader(docx)
            if not r:
                cells.append("   0/--    ")
                continue
            over = "*" if r[0][1] > RIGHT + 0.3 else " "
            cells.append("%d/%7.2f%s" % (len(r), r[0][1], over))
        print("  %-8s %-8s %s" % (j, p, "  ".join(cells)))
print("\n  * = the line's last glyph reaches past the column's right edge")
