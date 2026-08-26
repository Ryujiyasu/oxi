# -*- coding: utf-8 -*-
r"""Does the room a number format reserves take part in centring a cell?

`0_)` tells Excel to leave a blank the width of `)` after the number, so a
positive lines up with a negative that ends in one. Our formatter writes that
blank as a space and then centres the whole string, which puts the digits half
a blank left of centre: `procurement_contractor_list_02`'s company numbers sit
five pixels left of Excel's across 186 tiles of its picture.

So this asks Excel, one cell an arm: a number under five formats — no reserve,
a reserve after, a reserve before, both, and the two-section form the book
uses — against each of the three alignments, in a column wide enough that
nothing is squeezed. The answer is the ink's own start and end inside the
column, so no assumption about where the string is measured from comes into
it.

    python tools\metrics\_xlsx_reserved_align.py
    python tools\metrics\_xlsx_reserved_align.py --reuse
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SCRATCH = Path(r"C:\tmp\xlsx_reserved_align")
ROW_PT = 26.0
COLUMN = 22.0
VALUE = 12345
FORMATS = ["0", "0_)", "_(0", "_(0_)", "0_);(0)", "0_ ", "#,##0.0_ "]
ALIGNS = [("left", -4131), ("centre", -4108), ("right", -4152)]
FACES = [("ＭＳ 明朝", 14.0, True), ("Calibri", 11.0, False)]
WRAPS = [False, True]
ARMS = [(face, points, bold, code, name, how, wrap)
        for face, points, bold in FACES
        for code in FORMATS
        for name, how in ALIGNS
        for wrap in WRAPS]


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:C{len(ARMS) + 4}").Interior.Color = 0xFFFFFF
        sheet.Columns(2).ColumnWidth = COLUMN
        for at, (face, points, bold, code, _name, how, wrap) in enumerate(ARMS, start=2):
            cell = sheet.Cells(at, 2)
            cell.Value = VALUE
            try:
                cell.NumberFormat = code
            except Exception as trouble:
                print(f'  Excel refused {code!r}: {trouble}')
                raise
            cell.Font.Name = face
            cell.Font.Size = points
            cell.Font.Bold = bold
            cell.HorizontalAlignment = how
            cell.VerticalAlignment = -4160
            cell.WrapText = wrap
            sheet.Rows(at).RowHeight = ROW_PT
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(ARMS), 2)).CopyPicture(
                    Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.8)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def ours(made: Path):
    """Our picture, the column's edges, and the top of every arm's row.

    The row pitch must be READ, not assumed: `round(26pt * 96 / 72)` is 35 and
    the rows are 34, so by the fortieth arm the band being measured belongs to
    a different arm than the one it is labelled with, and the whole table slides
    without ever looking wrong.
    """
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
                          env={"OXI_XLSX_DUMP_COLUMNS": "1", "OXI_XLSX_DUMP_ROWS": "1",
                               **os.environ},
                          capture_output=True, text=True, encoding="utf-8")
    columns, at = {}, 0
    rows, down = {}, 0
    for line in (told.stdout or "").splitlines() + (told.stderr or "").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column":
            columns[int(parts[1])] = (at, at + int(parts[3]))
            at += int(parts[3])
        if len(parts) == 4 and parts[0] == "row":
            rows[int(parts[1])] = (down, down + int(parts[3]))
            down += int(parts[3])
    return np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140, columns, rows


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "reserved.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    mine, columns, rows = ours(made)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    wide = columns[1][1] - columns[1][0]
    left, right = columns[1]
    first = rows[2][0]
    print(f"  {VALUE} in a {wide}px column, {len(ARMS)} arms")
    print(f"  {'face':<10}{'format':<18}{'align':<8}"
          f"{'Excel ink':>16}{'Oxi ink':>14}   dx")
    agree = 0
    for at, (face, points, bold, code, name, _how, wrap) in enumerate(ARMS):
        top, foot = rows[at + 2]
        band = truth[top - first + 1:foot - first - 1, 1:wide - 1]
        lit = np.where(band.any(axis=0))[0]
        theirs = (int(lit.min()) + 1, int(lit.max()) + 1) if len(lit) else None
        band = mine[top + 1:foot - 1, left + 1:right - 1]
        lit = np.where(band.any(axis=0))[0]
        ours_at = (int(lit.min()) + 1, int(lit.max()) + 1) if len(lit) else None
        if theirs is None or ours_at is None:
            print(f"  {face:<10}{code:<18}{name:<8}{wrap!s:<7}  nothing to read")
            continue
        dx = ours_at[0] - theirs[0]
        agree += dx == 0 and ours_at[1] == theirs[1]
        print(f"  {face:<10}{code:<18}{name:<8}"
              f"{str(theirs):>16}{str(ours_at):>14}  {dx:>+3}"
              f"{'' if dx == 0 else '  <<'}")
    print(f"  {agree} of {len(ARMS)} arms agree")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
