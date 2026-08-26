# -*- coding: utf-8 -*-
r"""How tall is the box Excel draws for a note?

The corpus floor, `b6a3a84180c9_002`, drops a whole line of a cell comment:
Excel draws the note's second line, we do not. The reason is the box — Excel's
is 79 pixels where ours is 77, and the line needs those two.

The VML states the size in points (`width:175.5pt;height:59pt`), so the
question is what Excel does with that number. 59pt is 78.67px at 96dpi, which
rounds to 79 and truncates to 78 — neither is 77, and the sheet holds notes of
two different heights, so the arithmetic cannot be settled by reading the file.

This asks instead: one note an arm, its height set through COM, and the box
read straight off Excel's own picture. The border is what is measured — a note
is filled, so its edge is where the fill stops.

    python tools\metrics\_xlsx_note_box.py
    python tools\metrics\_xlsx_note_box.py --reuse
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
SCRATCH = Path(r"C:\tmp\xlsx_note_box")
HEIGHTS = [30.0, 40.0, 45.75, 50.0, 57.75, 58.0, 58.5, 59.0, 59.25, 60.0, 72.0, 82.5]
WIDE = 120.0
GAP = 3          # rows between one note and the next


def build() -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:H400").Interior.Color = 0xFFFFFF
        sheet.Columns(1).ColumnWidth = 3
        sheet.Columns(2).ColumnWidth = 30
        at = 2
        for tall in HEIGHTS:
            cell = sheet.Cells(at, 2)
            cell.ClearComments()
            note = cell.AddComment("x")
            note.Shape.Width = WIDE
            note.Shape.Height = tall
            note.Shape.Top = sheet.Cells(at, 2).Top
            note.Shape.Left = sheet.Cells(at, 2).Left
            note.Visible = True
            at += GAP + int(tall / 15) + 2
        rows = at + 4
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(rows, 8)).CopyPicture(
                    Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.9)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def boxes(picture: Image.Image) -> list[tuple[int, int]]:
    """Every note's top and bottom, found by its fill."""
    rgb = np.asarray(picture.convert("RGB")).astype(int)
    # Excel's note is #FFFFE1 unless told otherwise. The tolerance has to be
    # tight: white is only thirty away from it, so a loose one calls the whole
    # sheet a note.
    fill = (np.abs(rgb - np.array([255, 255, 225])).sum(axis=2) < 12)
    rows = np.where(fill.sum(axis=1) > 20)[0]
    out, start, last = [], None, None
    for y in rows:
        if start is None:
            start = last = y
            continue
        if y > last + 1:
            out.append((int(start), int(last)))
            start = y
        last = y
    if start is not None:
        out.append((int(start), int(last)))
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    if not args.reuse and not build():
        print("  Excel would not hand over a picture")
        return 1
    found = boxes(Image.open(SCRATCH / "excel.png"))
    print(f"  {len(HEIGHTS)} note(s) asked for, {len(found)} box(es) drawn")
    print(f"  {'asked pt':>9}{'96dpi px':>10}{'drawn px':>10}   what that is")
    for tall, held in zip(HEIGHTS, found):
        drawn = held[1] - held[0] + 1
        exact = tall * 96 / 72
        # The fill stops one inside the border on each side, so the box the
        # renderer is asked for is two more than the fill.
        note = []
        for name, value in (("round", round(exact)), ("floor", int(exact)),
                            ("ceil", -(-exact // 1))):
            if drawn == value:
                note.append(name)
            if drawn == value + 2:
                note.append(f"{name}+2")
        print(f"  {tall:>9}{exact:>10.2f}{drawn:>10}   {', '.join(note) or '—'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
