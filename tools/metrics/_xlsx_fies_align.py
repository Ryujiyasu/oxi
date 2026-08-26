# -*- coding: utf-8 -*-
r"""Is it the distributed alignment that lifts `fies_t2`'s labels a pixel?

Every distributed label in that book's worst column sits a pixel above ours,
at 19px rows and at 34px ones alike, while the right-aligned numbers beside
them land exactly. A scratch workbook holding the same text in the same face
at the same size and row height reproduces none of it, so either the alignment
is not the cause or the scratch is missing what carries it.

This asks the book itself. It opens a copy, photographs the label column, then
changes ONE thing — those cells' horizontal alignment — and photographs it
again. If the ink moves, the alignment carries it.

    python tools\metrics\_xlsx_fies_align.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
BOOK = REPO / "tools" / "golden-test" / "documents" / "xlsx" / "d1eb05860dd5_fies_t2.xlsx"
SCRATCH = Path(r"C:\tmp\xlsx_fies_align")
ROWS = (127, 145)
COLUMNS = "G:L"


def shot(sheet, excel, name: str) -> Image.Image | None:
    for _ in range(10):
        try:
            sheet.Activate()
            sheet.Range(f"G{ROWS[0]}:L{ROWS[1]}").CopyPicture(Appearance=1, Format=2)
        except Exception:
            time.sleep(0.6)
            continue
        time.sleep(0.8)
        held = ImageGrab.grabclipboard()
        if held is not None:
            held.save(SCRATCH / f"{name}.png")
            return held
    return None


def bands(picture: Image.Image) -> list[tuple[int, int]]:
    dark = np.asarray(picture.convert("L")) < 140
    rows = [r for r in range(dark.shape[0]) if dark[r].sum() >= 4]
    out, start, last = [], None, None
    for r in rows:
        if start is None:
            start = last = r
            continue
        if r > last + 1:
            out.append((start, last))
            start = r
        last = r
    if start is not None:
        out.append((start, last))
    return out


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(BOOK), ReadOnly=True)
    try:
        sheet = book.Worksheets(1)
        before = shot(sheet, excel, "before")
        if before is None:
            print("  Excel would not hand over a picture")
            return 1
        was = bands(before)
        # One thing only: the alignment of the label cells.
        sheet.Range(f"G{ROWS[0]}:G{ROWS[1]}").HorizontalAlignment = -4152   # right
        after = shot(sheet, excel, "after")
        now = bands(after)
        print(f"  {BOOK.name}  G{ROWS[0]}:G{ROWS[1]}, distributed -> right")
        print(f"  {len(was)} bands before, {len(now)} after")
        for at, (one, two) in enumerate(zip(was, now)):
            mark = "" if one[0] == two[0] else f"   moved {two[0] - one[0]:+d}"
            print(f"    band {at:>2}  {str(one):>12} -> {str(two):>12}{mark}")
        return 0
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
