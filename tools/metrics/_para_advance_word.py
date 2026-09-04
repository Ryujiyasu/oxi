# -*- coding: utf-8 -*-
"""Word's own y for a run of consecutive paragraphs, including the empty ones.

An empty paragraph draws no glyph, so the exported PDF cannot say where it sits
or how tall it is -- a gap between two visible lines could be one empty
paragraph or three. COM can: `Information(6)` reports the vertical position of a
collapsed range, empty paragraphs included.

Per CLAUDE.local.md the range MUST be collapsed at the paragraph start
(`doc.Range(rng.Start, rng.Start)`); the paragraph's own range reports its
ACTIVE END, which lands on the next page for a paragraph whose trailing mark
overflows.

    python _para_advance_word.py <docx> <substring> [count]

Prints the paragraph containing <substring> and the following <count>
paragraphs, with the advance from each to the next.
"""
import os
import sys

import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

wdActiveEndPageNumber = 3
wdVerticalPositionRelativeToPage = 6


def main():
    path = os.path.abspath(sys.argv[1])
    want = sys.argv[2]
    count = int(sys.argv[3]) if len(sys.argv) > 3 else 5
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    doc = app.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False)
    try:
        hit = None
        for i, p in enumerate(doc.Paragraphs, 1):
            if want in (p.Range.Text or ""):
                hit = i
                break
        if hit is None:
            print("not found:", want)
            return
        rows = []
        for i in range(hit, min(hit + count + 1, doc.Paragraphs.Count + 1)):
            p = doc.Paragraphs(i)
            st = doc.Range(p.Range.Start, p.Range.Start)
            rows.append((
                i,
                int(st.Information(wdActiveEndPageNumber)),
                float(st.Information(wdVerticalPositionRelativeToPage)),
                (p.Range.Text or "").rstrip("\r\x07"),
            ))
        print(f"{'para':>5} {'pg':>3} {'y':>8} {'advance':>8}  text")
        for k, (i, pg, y, t) in enumerate(rows):
            adv = "" if k + 1 >= len(rows) else (
                "(page)" if rows[k + 1][1] != pg else "%.2f" % (rows[k + 1][2] - y))
            print(f"{i:>5} {pg:>3} {y:8.2f} {adv:>8}  {t[:34]!r}")
    finally:
        doc.Close(False)
        app.Quit()


if __name__ == "__main__":
    main()
