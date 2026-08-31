# -*- coding: utf-8 -*-
"""Word truth for _pb_txbxfit_gen arms: which text-box lines survive, and where."""
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_txbxfit_gen import ARMS, EXTRA, LMARKS, OUT, MARKS  # noqa: E402

import fitz  # noqa: E402
import win32com.client  # noqa: E402

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


try:
    print("%-12s %6s %5s %4s %4s  %s" % ("arm", "boxH", "lines", "clip", "grid", "kept (mark@baseline)"))
    for tag in ARMS:
        height_pt, nlines, overflow, grid = ARMS[tag]
        p = os.path.join(OUT, tag + ".docx")
        pdf = p[:-5] + ".pdf"
        d = retry(lambda: word.Documents.Open(os.path.abspath(p), ReadOnly=True))
        try:
            retry(lambda: d.SaveAs2(os.path.abspath(pdf), FileFormat=17))
        finally:
            retry(lambda: d.Close(False))
        doc_ = fitz.open(pdf)
        kept = []
        body1 = None
        for pg in doc_:
            for b in pg.get_text("dict")["blocks"]:
                if b["type"] != 0:
                    continue
                for line in b["lines"]:
                    for s in line["spans"]:
                        ex = EXTRA.get(tag, ())
                        marks = LMARKS if (len(ex) > 5 and ex[5]) else MARKS
                        for mi, m in enumerate(marks[:nlines]):
                            if m in s["text"]:
                                kept.append("%s@%.2f" % (m, s["origin"][1]))
                        if "L01" in s["text"] and body1 is None:
                            body1 = s["origin"][1]
        doc_.close()
        print("%-12s %6.2f %5d %4s %4s  %-40s body_L01@%.2f"
              % (tag, height_pt, nlines, "clip" if overflow else "-",
                 "grid" if grid else "-", " ".join(kept), body1 or -1))
finally:
    word.Quit()
