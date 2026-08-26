# -*- coding: utf-8 -*-
r"""Which input puts `fies_t2`'s merged label a pixel higher than a scratch one?

The book's merged, distributed labels sit at offset 2 inside their row; a
scratch workbook holding the same text, face, size, row height, alignment and
merge puts them at 3 — the opposite of the difference against our renderer, so
some input is still unaccounted for. Two candidates survive: the book merges
from a column only 12px wide (G is 1.5 characters, H is 12.1), and its Normal
style is Terminal 14pt rather than the usual gothic.

So this adds them one at a time. Two workbooks — one with the ordinary Normal
style, one with the book's — each holding the same arms over two column
layouts. Whichever addition moves Excel's ink from 3 to 2 is the real input.

    python tools\metrics\_xlsx_fies_bisect.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
SCRATCH = Path(r"C:\tmp\xlsx_fies_bisect")
WORDS = "女性用洋服"
POINTS = 12.0
TALL = 14.25
# The book's own G..K, and a plain layout of the same total width.
BOOK_WIDTHS = [1.5, 12.09765625, 3.5, 1.59765625, 3.5]
PLAIN_WIDTHS = [sum(BOOK_WIDTHS) / 5] * 5
LAYOUTS = [("plain", PLAIN_WIDTHS), ("book", BOOK_WIDTHS)]
ALIGNS = [("distributed", -4117), ("right", -4152)]
ARMS = [(layout, name, how, merged)
        for layout, _widths in LAYOUTS
        for name, how in ALIGNS
        for merged in (True, False)]


def measure(normal: bool) -> list[tuple[int, int] | None]:
    """Every arm's ink band, in a workbook whose Normal style is the book's or not."""
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        if normal:
            book.Styles("Normal").Font.Name = "Terminal"
            book.Styles("Normal").Font.Size = 14
        sheet = book.Worksheets(1)
        sheet.Range("A1:L200").Interior.Color = 0xFFFFFF
        # Two layouts side by side: columns B..F and H..L.
        for at, width in enumerate(PLAIN_WIDTHS):
            sheet.Columns(2 + at).ColumnWidth = width
        for at, width in enumerate(BOOK_WIDTHS):
            sheet.Columns(8 + at).ColumnWidth = width
        sheet.Columns(7).ColumnWidth = 2.0
        for at, (layout, _name, how, merged) in enumerate(ARMS, start=2):
            first = 2 if layout == "plain" else 8
            cell = sheet.Cells(at, first)
            cell.Value = WORDS
            cell.Font.Name = "ＭＳ 明朝"
            cell.Font.Size = POINTS
            cell.HorizontalAlignment = how
            if merged:
                sheet.Range(sheet.Cells(at, first), sheet.Cells(at, first + 4)).Merge()
            sheet.Rows(at).RowHeight = TALL
        # One row to a picture. Slicing a tall picture into bands by an
        # assumed pitch is how the first reading of this came out with the
        # sign reversed: `CopyPicture` frames the range, and the frame is not
        # where the assumed origin puts it.
        out = []
        for at, _arm in enumerate(ARMS, start=2):
            held = None
            for _ in range(10):
                try:
                    sheet.Activate()
                    sheet.Range(sheet.Cells(at, 2), sheet.Cells(at, 12)).CopyPicture(
                        Appearance=1, Format=2)
                except Exception:
                    time.sleep(0.6)
                    continue
                time.sleep(0.7)
                held = ImageGrab.grabclipboard()
                if held is not None:
                    break
            if held is None:
                out.append(None)
                continue
            held.save(SCRATCH / f"normal_{normal}_arm{at}.png")
            dark = np.asarray(held.convert("L")) < 140
            rows = [r for r in range(dark.shape[0]) if dark[r].sum() >= 4]
            out.append((rows[0], rows[-1]) if rows else None)
        return out
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def main() -> int:
    plain = measure(normal=False)
    booked = measure(normal=True)
    print(f"  {WORDS!r}, ＭＳ 明朝 {POINTS}pt, row {TALL} ({round(TALL * 96 / 72)}px)")
    print(f"  the book's own cell reads offset 2; a scratch one reads 3")
    print(f"  {'columns':<8}{'align':<13}{'merged':<8}"
          f"{'Normal ordinary':>17}{'Normal Terminal 14':>20}")
    for at, (layout, name, _how, merged) in enumerate(ARMS):
        print(f"  {layout:<8}{name:<13}{merged!s:<8}"
              f"{str(plain[at]):>17}{str(booked[at]):>20}"
              f"{'   <<' if plain[at] != booked[at] else ''}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
