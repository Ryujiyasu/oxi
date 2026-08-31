# -*- coding: utf-8 -*-
"""Word truth for _pb_negtop_gen: first body baseline vs the declared top margin."""
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_negtop_gen import ARMS, OUT  # noqa: E402

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
    print("%-14s %8s %7s %4s %5s | %9s %9s %9s"
          % ("arm", "top_tw", "top_pt", "hdrL", "grid", "L01_base", "H01_base", "L02-L01"))
    for tag in ARMS:
        a = ARMS[tag]
        top_tw, header_tw, header_lines, grid = a[:4]
        bottom_tw = a[4] if len(a) > 4 else -284
        p = os.path.join(OUT, tag + ".docx")
        pdf = p[:-5] + ".pdf"
        d = retry(lambda: word.Documents.Open(os.path.abspath(p), ReadOnly=True))
        try:
            retry(lambda: d.SaveAs2(os.path.abspath(pdf), FileFormat=17))
        finally:
            retry(lambda: d.Close(False))
        doc_ = fitz.open(pdf)
        ys = {}
        p1_last = None
        p1_count = 0
        for pgi, pg in enumerate(doc_):
            if pgi == 0:
                for b in pg.get_text("dict")["blocks"]:
                    if b["type"] != 0:
                        continue
                    for line in b["lines"]:
                        for s in line["spans"]:
                            t = s["text"].strip()
                            if t.startswith("L") and t[1:].isdigit():
                                p1_count += 1
                                if p1_last is None or s["origin"][1] > p1_last[0]:
                                    p1_last = (s["origin"][1], t)
        for pg in doc_:
            for b in pg.get_text("dict")["blocks"]:
                if b["type"] != 0:
                    continue
                for line in b["lines"]:
                    for s in line["spans"]:
                        t = s["text"].strip()
                        if t in ("L01", "L02", "H01") and t not in ys:
                            ys[t] = s["origin"][1]
            break  # page 1 only
        doc_.close()
        l1 = ys.get("L01")
        l2 = ys.get("L02")
        print("%-14s %8d %7.2f %4d %5s | %9s %9s %9s | bot=%6d p1_lines=%2d p1_last=%s@%s"
              % (tag, top_tw, top_tw / 20.0, header_lines, "grid" if grid else "-",
                 ("%.2f" % l1) if l1 else "-",
                 ("%.2f" % ys["H01"]) if "H01" in ys else "-",
                 ("%.2f" % (l2 - l1)) if (l1 and l2) else "-",
                 bottom_tw, p1_count,
                 p1_last[1] if p1_last else "-",
                 ("%.2f" % p1_last[0]) if p1_last else "-"))
finally:
    word.Quit()
