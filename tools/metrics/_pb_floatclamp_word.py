# -*- coding: utf-8 -*-
"""Word truth for _pb_floatclamp_gen arms: WHERE does Word draw the box?

Reads the PDF origin of the box's single marker glyph.  Reported next to the
two predictions:

    raw    = the anchor as written        (no clamp)
    clamp  = (page - box).max(0.0)        (what Oxi does today)

`lIns` is 7.2pt in every arm, so glyph_x - 7.2 is the resolved box left.
"""
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_floatclamp_gen import ARMS, MARK, OUT  # noqa: E402

import fitz  # noqa: E402
import win32com.client  # noqa: E402

LINS = 7.2
TINS = 3.6

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0


def retry(fn, tries=10):
    for i in range(tries):
        try:
            return fn()
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(1.5)


def col_left(page):
    """Left edge of column 1 in points."""
    return page["mar"] / 20.0


try:
    print("%-10s %8s %8s %8s | %8s %8s | %8s %8s %5s"
          % ("arm", "pgW", "boxW", "off", "raw_x", "clamp_x", "got_x", "got_y", "page"))
    for tag in ARMS:
        page, w, h, hrel, hoff, vrel, voff, lead = ARMS[tag]
        pw = page["pw"] / 20.0
        ph = page["ph"] / 20.0
        if page["orient"] == "landscape":
            pw, ph = max(pw, ph), min(pw, ph)
        ref_left = 0.0 if hrel == "page" else col_left(page)
        raw_x = ref_left + hoff
        clamp_x = (pw - w) if (raw_x + w > pw) else raw_x
        if clamp_x < 0.0:
            clamp_x = 0.0

        p = os.path.join(OUT, tag + ".docx")
        pdf = p[:-5] + ".pdf"
        d = retry(lambda: word.Documents.Open(os.path.abspath(p), ReadOnly=True))
        try:
            retry(lambda: d.SaveAs2(os.path.abspath(pdf), FileFormat=17))
        finally:
            retry(lambda: d.Close(False))

        got = []
        doc_ = fitz.open(pdf)
        for pi, pg in enumerate(doc_):
            for b in pg.get_text("dict")["blocks"]:
                if b["type"] != 0:
                    continue
                for line in b["lines"]:
                    for s in line["spans"]:
                        if MARK in s["text"]:
                            got.append((pi + 1, s["origin"][0], s["origin"][1]))
        doc_.close()
        if got:
            pi, gx, gy = got[0]
            print("%-10s %8.2f %8.2f %8.2f | %8.2f %8.2f | %8.2f %8.2f %5d   left=%.2f"
                  % (tag, pw, w, hoff, raw_x + LINS, clamp_x + LINS, gx, gy, pi, gx - LINS))
        else:
            print("%-10s %8.2f %8.2f %8.2f | %8.2f %8.2f | %8s %8s %5s"
                  % (tag, pw, w, hoff, raw_x + LINS, clamp_x + LINS, "-", "-", "-"))
finally:
    word.Quit()
