# -*- coding: utf-8 -*-
r"""Where does a WRAPPED, merged, distributed block sit when it is centred?

The merged pixel of leading is taken from the foot for a distributed line and
from the block for every other alignment — measured over 90 arms, but all of
them one line and none of them wrapped. Applying it to centred cells moved five
workbooks the wrong way, and every one of those five holds merged distributed
cells that are centred, while the two it moved the right way hold bottom ones.

So this asks the case the sweep never covered: two and three lines, wrapped,
merged, distributed, centred, over a range of row heights.

    python tools\metrics\_xlsx_spread_wrapped.py
    python tools\metrics\_xlsx_spread_wrapped.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_spread_wrapped")
POINTS = 12.0
COLUMN = 9.0
TEXTS = [("one", "女性用洋服"),
         ("two", "女性用洋服とシャツセーター類"),
         ("three", "女性用洋服とシャツセーター類および子供用の洋服")]
UPRIGHT = [("centre", -4108), ("bottom", -4107)]
HEIGHTS = [30.0, 36.0, 42.0, 48.0]
ARMS = [(words, name, tall, sit, sits, merged)
        for name, words in TEXTS
        for sit, sits in UPRIGHT
        for tall in HEIGHTS
        for merged in (True, False)]


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:F{len(ARMS) + 4}").Interior.Color = 0xFFFFFF
        for column in (2, 3, 4):
            sheet.Columns(column).ColumnWidth = COLUMN
        for at, (words, _name, tall, _sit, sits, merged) in enumerate(ARMS, start=2):
            cell = sheet.Cells(at, 2)
            cell.Value = words
            cell.Font.Name = "ＭＳ 明朝"
            cell.Font.Size = POINTS
            cell.HorizontalAlignment = -4117            # distributed
            cell.VerticalAlignment = sits
            cell.WrapText = True
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


def ink(band: np.ndarray):
    rows = [r for r in range(band.shape[0]) if band[r].sum() >= 4]
    return (rows[0], rows[-1]) if rows else None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "wrapped.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    mine, columns, rows = ours(made)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    wide = columns[4][1] - columns[2][0]
    left, right = columns[2][0], columns[4][1]
    first = rows[2][0]
    print(f"  distributed, wrapped, ＭＳ 明朝 {POINTS}pt over {wide}px")
    print(f"  {'lines':<7}{'sits':<8}{'row':>5}{'px':>4}{'merged':>8}"
          f"{'Excel ink':>13}{'Oxi ink':>13}   dy")
    agree = 0
    for at, (_words, name, tall, sit, _sits, merged) in enumerate(ARMS):
        top, foot = rows[at + 2]
        theirs = ink(truth[top - first:foot - first, 1:wide - 1])
        ours_at = ink(mine[top:foot, left + 1:right - 1])
        if theirs is None or ours_at is None:
            print(f"  {name:<7}{sit:<8}{tall:>5}{foot - top:>4}{merged!s:>8}  nothing to read")
            continue
        agree += theirs == ours_at
        print(f"  {name:<7}{sit:<8}{tall:>5}{foot - top:>4}{merged!s:>8}"
              f"{str(theirs):>13}{str(ours_at):>13}  {ours_at[0] - theirs[0]:>+3}"
              f"{'' if theirs == ours_at else '  <<'}")
    print(f"  {agree} of {len(ARMS)} arms agree")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
