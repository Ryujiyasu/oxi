# -*- coding: utf-8 -*-
r"""Does a cell's horizontal alignment move its baseline?

`fies_t2`'s label column is the worst band in its picture: every distributed
label sits a pixel below Excel's, while the right-aligned numbers beside it —
same face, same size, same row, same box — land exactly. Our two paths draw at
the same baseline, so the pixel is Excel's.

The rows there are tight: 14.25pt is 19px and ＭＳ 明朝 12pt asks for about
that, so the question is really two. This sweeps the row height across the
size's own line height AND the six alignments, and reads the ink's top and
foot. A step that follows the alignment is one answer; a step that follows the
height is another.

    python tools\metrics\_xlsx_align_baseline.py
    python tools\metrics\_xlsx_align_baseline.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_align_baseline")
WORDS = "女性用洋服"
COLUMN = 26.0
ALIGNS = [("left", -4131), ("centre", -4108), ("right", -4152),
          ("distributed", -4117), ("justify", -4130), ("fill", 5)]
HEIGHTS = [13.5, 14.25, 15.0, 15.75, 16.5, 18.0, 20.25]
MERGED = [False, True]
ARMS = [(points, tall, name, how, merged)
        for points in (12.0,)
        for tall in HEIGHTS
        for name, how in ALIGNS
        for merged in MERGED]


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
        for _held in (3, 4):
            sheet.Columns(_held).ColumnWidth = COLUMN / 3
        for at, (points, tall, _name, how, merged) in enumerate(ARMS, start=2):
            cell = sheet.Cells(at, 2)
            cell.Value = WORDS
            cell.Font.Name = "ＭＳ 明朝"
            cell.Font.Size = points
            cell.HorizontalAlignment = how
            if merged:
                sheet.Range(sheet.Cells(at, 2), sheet.Cells(at, 4)).Merge()
            sheet.Rows(at).RowHeight = tall
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(ARMS), 4)).CopyPicture(
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
    columns, at, rows, down = {}, 0, {}, 0
    for line in (told.stdout or "").splitlines() + (told.stderr or "").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column":
            columns[int(parts[1])] = (at, at + int(parts[3]))
            at += int(parts[3])
        if len(parts) == 4 and parts[0] == "row":
            rows[int(parts[1])] = (down, down + int(parts[3]))
            down += int(parts[3])
    return np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140, columns, rows


def ink(band: np.ndarray) -> tuple[int, int] | None:
    """The band's first and last row that carry more than a rule's worth."""
    rows = [r for r in range(band.shape[0]) if band[r].sum() >= 4]
    return (rows[0], rows[-1]) if rows else None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "align.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    mine, columns, rows = ours(made)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    wide = columns[3][1] - columns[1][0]
    left, right = columns[1][0], columns[3][1]
    first = rows[2][0]
    print(f"  {WORDS!r} in ＭＳ 明朝, a {wide}px column")
    print(f"  {'pt':>5}{'row':>7}{'px':>4}  {'align':<12}"
          f"{'Excel ink':>12}{'Oxi ink':>12}   dy")
    agree = 0
    for at, (points, tall, name, _how, merged) in enumerate(ARMS):
        top, foot = rows[at + 2]
        theirs = ink(truth[top - first:foot - first, 1:wide - 1])
        ours_at = ink(mine[top:foot, left + 1:right - 1])
        if theirs is None or ours_at is None:
            print(f"  {points:>5}{tall:>7}{foot - top:>4}  {name:<12}{merged!s:<8}  nothing to read")
            continue
        dy = ours_at[0] - theirs[0]
        agree += theirs == ours_at
        print(f"  {points:>5}{tall:>7}{foot - top:>4}  {name:<12}"
              f"{str(theirs):>12}{str(ours_at):>12}  {dy:>+3}"
              f"{'' if theirs == ours_at else '  <<'}")
    print(f"  {agree} of {len(ARMS)} arms agree")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
