# -*- coding: utf-8 -*-
"""COM-measure harassmanual per-para top Y (R30 collapsed-start) for paras 1-24,
diff consecutive -> para heights -> line counts (lineRule=exact 18.5pt).
Compare to Oxi layout dump line counts. Localizes the wrap-width drift."""
import sys, io, json, os
import win32com.client as win32
DOCX = os.path.abspath("tools/golden-test/documents/docx/harassmanual_001466344.docx")
LH = 18.5  # 370tw exact
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
rows = []
try:
    n = doc.Paragraphs.Count
    for i in range(1, min(n, 26)+1):
        rng = doc.Paragraphs(i).Range
        col = doc.Range(rng.Start, rng.Start)
        y = col.Information(6)   # wdVerticalPositionRelativeToPage
        pg = col.Information(3)  # wdActiveEndPageNumber
        txt = rng.Text.strip()[:22]
        rows.append((i, pg, y, txt))
finally:
    doc.Close(False); word.Quit()
print("idx pg     y    dY  lines  text")
prev_y = None; prev_pg = None
for (i, pg, y, txt) in rows:
    dy = ""; ln = ""
    if prev_y is not None and pg == prev_pg:
        d = y - prev_y; dy = f"{d:6.1f}"; ln = f"{d/LH:4.2f}"
    elif prev_y is not None and pg != prev_pg:
        dy = "  PAGE"; ln = "  -->"
    print(f"{i:3d} p{pg}  {y:6.1f} {dy} {ln}  {txt!r}")
    prev_y = y; prev_pg = pg
