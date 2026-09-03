# -*- coding: utf-8 -*-
"""Full text + line count of every paragraph that starts on one page.

The column-width probe reads a paragraph's start x and the x of the characters
it owns; matching those against the truth PDF's columns needs the paragraph's
WHOLE text, not a preview, because two neighbouring paragraphs in this corpus
read as one sentence (a classical line and its modern gloss) and a column can be
attributed to the wrong one from a 20-character prefix.

    python _vert_p7_paras.py <docx> <page>
"""
import os
import sys

import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

wdActiveEndPageNumber = 3
wdHorizontalPositionRelativeToPage = 5
wdStatisticLines = 1


def main():
    path, page = os.path.abspath(sys.argv[1]), int(sys.argv[2])
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    doc = app.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False)
    try:
        for i, p in enumerate(doc.Paragraphs, 1):
            rng = p.Range
            st = doc.Range(rng.Start, rng.Start)
            if int(st.Information(wdActiveEndPageNumber)) != page:
                continue
            txt = rng.Text.rstrip("\r\x07")
            x = float(st.Information(wdHorizontalPositionRelativeToPage))
            lines = int(rng.ComputeStatistics(wdStatisticLines))
            print(f"{i:>4} x={x:8.2f} lines={lines:>2} len={len(txt):>3}  {txt!r}")
    finally:
        doc.Close(False)
        app.Quit()


if __name__ == "__main__":
    main()
