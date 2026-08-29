# -*- coding: utf-8 -*-
"""Read the exact-line baseline probe. Usage: _pb_exactbase_read.py word|oxi

Prints, per arm, the FIRST line's baseline measured from the page's top margin
(70.90pt), and the line-to-line pitch. Word's baseline comes from the PDF span
origin (exact); Oxi's is recovered from the rendered INK, so the two are
comparable without any box-vs-ink convention.
"""
import os, sys, json, subprocess, tempfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_exactbase_gen import ARMS, OUT

TOP = 1418 / 20.0
DPI = 300
REND = os.path.abspath("tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe")


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
    out = []
    for blk in doc[0].get_text("dict")["blocks"]:
        for l in blk.get("lines", []):
            for s in l["spans"]:
                if s["text"].strip().startswith("L"):
                    out.append((s["origin"][1], s["bbox"][1]))
                    break
    out.sort()
    return out


def oxi_rows(docx):
    """Ink top of each line band, from the render (no convention involved)."""
    import numpy as np
    import PIL.Image as I
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([REND, docx, os.path.join(tmp, "p"), str(DPI)],
                       capture_output=True, timeout=600)
        p = os.path.join(tmp, "p_p1.png")
        if not os.path.exists(p):
            return []
        g = np.asarray(I.open(p).convert("L"), dtype=np.float32)
    ink = (255.0 - g)
    ink[ink < 40] = 0.0
    prof = ink.sum(axis=1)
    on = prof > prof.max() * 0.02
    bands, i = [], 0
    while i < len(on):
        if on[i]:
            j = i
            while j < len(on) and on[j]:
                j += 1
            bands.append((i * 72.0 / DPI, j * 72.0 / DPI))
            i = j
        else:
            i += 1
    return bands


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
print("%s   first-line position relative to the top margin (%.2fpt)\n" % (mode.upper(), TOP))
print("  arm               line_pt  first_ink-TOP   pitch")
for tag, face, sz, line in ARMS:
    docx = os.path.join(OUT, tag + ".docx")
    if not os.path.exists(docx):
        print("  %-18s MISSING" % tag)
        continue
    if mode == "word":
        r = word_rows(docx)
        if len(r) < 2:
            print("  %-18s (no lines)" % tag)
            continue
        first_ink = r[0][1] - TOP
        pitch = r[1][0] - r[0][0]
        extra = "  baseline-TOP=%6.2f" % (r[0][0] - TOP)
    else:
        b = oxi_rows(docx)
        if len(b) < 2:
            print("  %-18s (no bands)" % tag)
            continue
        first_ink = b[0][0] - TOP
        pitch = b[1][0] - b[0][0]
        extra = ""
    print("  %-18s %6.2f   %8.2f     %6.2f%s" % (tag, line / 20.0, first_ink, pitch, extra))
