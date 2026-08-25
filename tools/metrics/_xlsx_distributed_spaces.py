# -*- coding: utf-8 -*-
r"""Do the spaces round a distributed line take a share of the spread?

`fies_t2`'s headings are 518 cells of `horizontal="distributed"`, and many of
them hold a line with trailing spaces — `"有業人員  "`. We give every space a
piece of its own, so the visible glyphs are packed into the left of the cell
and the spread comes out 36 pixels where Excel's is 55. The question is
whether Excel drops them, and whether it treats the two ends alike.

Each arm is a cell of its own in the same column, so the width is shared and
the spreads are comparable: the reading is where the FIRST and LAST ink land
inside the cell.

    python tools\metrics\_xlsx_distributed_spaces.py
    python tools\metrics\_xlsx_distributed_spaces.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_distributed_spaces")
FACE, POINTS = "ＭＳ 明朝", 9.0
COLUMN = 24.0
ROW_PT = 18.0
ARMS = [
    "有業人員",
    "有業人員 ",
    "有業人員  ",
    " 有業人員",
    "  有業人員",
    "  有業人員  ",
    "有業 人員",
    "有業　人員",          # an ideographic space, which is not ASCII
    "AB CD",
    "AB CD  ",
]


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:D{len(ARMS) + 4}").Interior.Color = 0xFFFFFF
        sheet.Columns(2).ColumnWidth = COLUMN
        for at, words in enumerate(ARMS, start=2):
            cell = sheet.Cells(at, 2)
            # A leading space would be eaten by Excel's own parsing of a typed
            # value, so the text is set as a formula-free string outright.
            cell.NumberFormat = "@"
            cell.Value = words
            cell.Font.Name = FACE
            cell.Font.Size = POINTS
            cell.HorizontalAlignment = -4117        # xlDistributed
            cell.VerticalAlignment = -4108
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
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
                          env={"OXI_XLSX_DUMP_COLUMNS": "1", "OXI_XLSX_DUMP_ROWS": "1",
                               **os.environ},
                          capture_output=True, text=True, encoding="utf-8")
    columns, rows, at, down = {}, {}, 0, 0
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
    made = SCRATCH / "distributed.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    mine, columns, rows = ours(made)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    left, right = columns[1]
    wide = right - left
    print(f"  {FACE} {POINTS}pt distributed in a {wide}px cell")
    print("  text              |  ink: Excel        Oxi")
    walked = 0
    for at, words in enumerate(ARMS):
        top, foot = rows[at + 2]
        tall = foot - top
        band = truth[walked + 1:walked + tall - 1, 1:wide - 1]
        lit = np.where(band.any(axis=0))[0]
        theirs = (int(lit.min()) + 1, int(lit.max()) + 1) if len(lit) else None
        band = mine[top + 1:foot - 1, left + 1:right - 1]
        lit = np.where(band.any(axis=0))[0]
        ours_at = (int(lit.min()) + 1, int(lit.max()) + 1) if len(lit) else None
        walked += tall
        if theirs is None or ours_at is None:
            print(f"  {words!r:<18}|  nothing to read")
            continue
        print(f"  {words!r:<18}|  {theirs[0]:>3}-{theirs[1]:<4} {ours_at[0]:>8}-{ours_at[1]:<4}"
              f"  {'' if theirs == ours_at else '<<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
