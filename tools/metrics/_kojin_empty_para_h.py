# -*- coding: utf-8 -*-
"""Measure Word's rendered line height for the 14pt empty-paragraph run in kojin
(wi ~587-617), to determine whether a 14pt empty para in a type=lines linePitch=360
docGrid occupies 1 cell (18pt) or 2 cells (36pt). Information(6) collapsed-start.
"""
import os, sys
import win32com.client as win32

DOCX = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",
    "golden-test", "documents", "docx", "kojin_000505813.docx"))

word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        lo = int(sys.argv[1]) if len(sys.argv) > 1 else 583
        hi = int(sys.argv[2]) if len(sys.argv) > 2 else 622
        prev_y = None
        prev_pg = None
        for pi in range(lo, hi + 1):
            p = doc.Paragraphs(pi)
            rng = p.Range
            raw = rng.Text or ""
            sr = doc.Range(rng.Start, rng.Start)
            pg = sr.Information(3)
            y = sr.Information(6)  # pts from page top
            # font size of the paragraph
            try:
                sz = p.Range.Font.Size
            except Exception:
                sz = None
            gap = ""
            if prev_y is not None and prev_pg == pg:
                gap = f"gap={y - prev_y:.2f}"
            txt = raw.replace("\r", "").replace("\x07", "")[:18]
            print(f"wi={pi:3d} pg={pg} y={y:7.2f} sz={sz} {gap:14s} {txt!r}")
            prev_y = y
            prev_pg = pg
    finally:
        doc.Close(False)
finally:
    word.Quit()
