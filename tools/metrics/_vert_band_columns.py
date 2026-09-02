"""Per-COLUMN occupancy of a vertical multi-column page.

Walks characters of the paragraphs on one page and records, for each distinct
(y_band, x_column), the first character that lands there. That answers "which
columns of which band does each section actually occupy", which paragraph-start
positions alone cannot (a paragraph spanning 3 columns reports only its first).
"""
import json, os, sys
import win32com.client as win32

wdActiveEndSectionNumber = 2
wdActiveEndPageNumber = 3
wdHorizontalPositionRelativeToPage = 5
wdVerticalPositionRelativeToPage = 6


def run(path, page, out):
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    doc = app.Documents.Open(os.path.abspath(path), ReadOnly=True, AddToRecentFiles=False)
    try:
        cells = []          # (x, y, sec, para_i, char)
        for i, p in enumerate(doc.Paragraphs, 1):
            rng = p.Range
            st = doc.Range(rng.Start, rng.Start)
            if int(st.Information(wdActiveEndPageNumber)) != page:
                continue
            sec = int(st.Information(wdActiveEndSectionNumber))
            seen = set()
            for pos in range(rng.Start, rng.End):
                c = doc.Range(pos, pos + 1)
                txt = c.Text or ""
                if txt in ("\r", "\x07", ""):
                    continue
                cr = doc.Range(pos, pos)
                x = round(float(cr.Information(wdHorizontalPositionRelativeToPage)), 2)
                y = round(float(cr.Information(wdVerticalPositionRelativeToPage)), 2)
                if x in seen:
                    continue
                seen.add(x)
                cells.append({"x": x, "y": y, "sec": sec, "para": i, "ch": txt})
        json.dump({"page": page, "cells": cells}, open(out, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
        print("wrote", out, len(cells), "column-entries")
    finally:
        doc.Close(False)
        app.Quit()


if __name__ == "__main__":
    run(sys.argv[1], int(sys.argv[2]), sys.argv[3])
