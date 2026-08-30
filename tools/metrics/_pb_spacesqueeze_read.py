# -*- coding: utf-8 -*-
"""Read the space-squeeze probe. Usage: _pb_spacesqueeze_read.py word

For each (size, jc, word-count) group, walk L upward and report the first line's
space advance and how many lines the paragraph took. The BOUNDARY arm -- the
largest L still on one line -- carries the tightest space Word was willing to
set, i.e. the squeeze limit.
"""
import os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_spacesqueeze_gen import ARMS, OUT, SHAPES, LS, SIZES, JC

COL = 451.32


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
    # ★`dict` carries no per-character data -- only `rawdict` does, and the
    # space advance is the whole point here.
    for blk in doc[0].get_text("rawdict")["blocks"]:
        for l in blk.get("lines", []):
            chs = [c for s in l["spans"] for c in s.get("chars", [])]
            txt = "".join(c["c"] for c in chs)
            if not txt.strip() or txt.strip().startswith("AFTER"):
                continue
            adv = [chs[i + 1]["origin"][0] - chs[i]["origin"][0]
                   for i in range(len(chs) - 1) if chs[i]["c"] == " "]
            ink = chs[-1]["bbox"][2] - chs[0]["origin"][0]
            out.append((round(l["bbox"][1], 2), adv, ink))
    out.sort()
    return out


print("  arm group            L: lines  space(mean)  ink     [* = last that fits]")
for s in SIZES:
    for j in JC:
        for w in SHAPES:
            cells = []
            prev_fit = None
            for L in LS:
                docx = os.path.join(OUT, "%s_%s_%s_L%d.docx" % (s, j, w, L))
                if not os.path.exists(docx):
                    continue
                r = word_rows(docx)
                if not r:
                    continue
                n, adv, ink = len(r), r[0][1], r[0][2]
                sp = (sum(adv) / len(adv)) if adv else 0.0
                cells.append((L, n, sp, ink))
            fits = [c for c in cells if c[1] == 1]
            last = fits[-1][0] if fits else None
            print("  %-6s %-5s %-6s" % (s, j, w))
            for L, n, sp, ink in cells:
                mark = "*" if L == last else " "
                print("      L=%-3d lines=%d  space=%6.3f  ink=%7.2f %s" % (L, n, sp, ink, mark))
