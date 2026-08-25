# -*- coding: utf-8 -*-
r"""Where does a MIXED stack sit as the column widens?

`_xlsx_stack_centre.py` asks this at one width (53px) and reads that a stack
holding a turned mark is centred on the turned box — the standing character a
pixel left of the em's centre, the turned one a pixel right. Implementing that
gains `r03_syukei2` 0.0004 and costs `data_B22` 0.0046, whose stacks are the
same mixed kind but sit in cells **17 pixels wide — the turned box exactly** —
where Excel centres on the em instead.

So one rule cannot be both, and what is missing is the width. This sweeps the
column a pixel at a time with the same mixed stack in it and reads where each
character lands from the cell's own left edge, Excel's beside ours.

    python tools\metrics\_xlsx_stack_width.py
    python tools\metrics\_xlsx_stack_width.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_stack_width")
FACE, SIZE = "ＭＳ 明朝", 11.0
STACK = "相（"          # one standing, one turned
ROW_PT = 45.0
# A digit of the standard font is 7 pixels, so an eighth of a character is
# under a pixel: this walks the column's pixel width one at a time from the
# turned box (17px) up past `_xlsx_stack_centre.py`'s own 53.
WIDTHS = [1.0 + step / 8.0 for step in range(0, 40)]


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:BZ8").Interior.Color = 0xFFFFFF
        for at, width in enumerate(WIDTHS, start=2):
            sheet.Columns(at).ColumnWidth = width
            sheet.Columns(at).NumberFormat = "@"
            cell = sheet.Cells(2, at)
            cell.Value = STACK
            cell.Font.Name = FACE
            cell.Font.Size = SIZE
            cell.Orientation = -4166          # xlVertical
            cell.VerticalAlignment = -4160    # xlTop
            cell.HorizontalAlignment = -4108  # xlCenter
        sheet.Rows(2).RowHeight = ROW_PT
        book.SaveAs(str(made), FileFormat=51)
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(2, 1 + len(WIDTHS)))
        for _ in range(12):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(0.8)
            grabbed = ImageGrab.grabclipboard()
            if grabbed is not None:
                grabbed.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def ours(made: Path):
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
                          env={"OXI_XLSX_DUMP_COLUMNS": "1", **os.environ},
                          capture_output=True, text=True, encoding="utf-8")
    columns, at = {}, 0
    for line in (told.stdout or "").splitlines() + (told.stderr or "").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column":
            columns[int(parts[1])] = (at, at + int(parts[3]))
            at += int(parts[3])
    return np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140, columns


def first_ink(band: np.ndarray) -> int | None:
    lit = np.where(band.any(axis=0))[0]
    return int(lit.min()) if len(lit) else None


def per_line(band: np.ndarray) -> list[int]:
    """Each stacked character's own first ink, read in its own row band.

    Reading the cell as a whole gives the MINIMUM over both characters, and in
    a narrow cell the standing one and the turned one overlap — so a stack put
    at 0 and 2 reads the same as one put at 1 and 1. The two characters are on
    different LINES, so the answer is one band each.
    """
    rows = np.where(band.any(axis=1))[0]
    if not len(rows):
        return []
    groups, start, last = [], rows[0], rows[0]
    for y in rows[1:]:
        if y - last > 2:
            groups.append((start, last))
            start = y
        last = y
    groups.append((start, last))
    out = []
    for top, foot in groups:
        lit = np.where(band[top:foot + 1].any(axis=0))[0]
        if len(lit):
            out.append(int(lit.min()))
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "stackwidth.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    mine, columns = ours(made)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    print(f"  {FACE} {SIZE}pt, stack {STACK!r} (one standing, one turned)")
    print("  cell px | each character's own first ink, and the gap between them")
    walked = 0
    for at in range(len(WIDTHS)):
        if at + 1 not in columns:
            continue
        left, right = columns[at + 1]
        wide = right - left
        theirs = per_line(truth[:, walked + 1:walked + wide - 1])
        ours_at = per_line(mine[:, left + 1:left + wide - 1])
        walked += wide
        if len(theirs) != 2 or len(ours_at) != 2:
            print(f"  {wide:>7} | read {len(theirs)}/{len(ours_at)} lines, not two")
            continue
        print(f"  {wide:>7} | 相 E {theirs[0] + 1:>2} O {ours_at[0] + 1:<2}"
              f"   （ E {theirs[1] + 1:>2} O {ours_at[1] + 1:<2}"
              f"   gap E {theirs[1] - theirs[0]:>2} O {ours_at[1] - ours_at[0]:<2}"
              f"  {'' if theirs == ours_at else '<<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
