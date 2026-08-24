"""Excel and Oxi on the same stacked cells, offset by offset.

`_xlsx_stack_center.py` measured Excel alone: a stacked block's height is
`14n - 4` for ＭＳ 明朝 at 8pt, constant across row heights, and the leftover
half is rounded UP — an odd row puts the spare pixel above the block. This
builds the same sheet, saves it, and renders it BOTH ways, so the two offsets
can be read off one picture each and subtracted. Guessing which of the block
or the rounding is wrong from Excel's numbers alone cannot separate them.

Run: python tools/metrics/_xlsx_stack_ab.py
"""

from __future__ import annotations

import os
import re
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
SCRATCH = Path(r"C:\tmp\xlsx_stack_ab")
ARMS = [("ＭＳ 明朝", 8.0), ("ＭＳ ゴシック", 8.0), ("ＭＳ Ｐゴシック", 11.0),
        ("ＭＳ 明朝", 11.0), ("游ゴシック", 9.0), ("Meiryo UI", 9.0)]
WORDS = ["続柄", "性別番", "就業期間中"]
HIGHS = [80.0, 90.0, 105.0, 117.9, 118.5]


def offsets(grey: np.ndarray, tops: list[int], highs: list[int], left: int, wide: int):
    out = []
    for top, high in zip(tops, highs):
        block = grey[top:top + high, left:left + wide]
        lit = np.where((block < 140).any(axis=1))[0]
        out.append(int(lit.min()) if len(lit) else None)
    return out


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "stack.xlsx"
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    tops, highs = [], []
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:F40").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = 4.0
        at = 2
        for face, size in ARMS:
          for words in WORDS:
            for high in HIGHS:
                cell = sheet.Cells(at, 2)
                cell.Value = words
                cell.Font.Name = face
                cell.Font.Size = size
                cell.Orientation = -4166           # xlVertical
                cell.VerticalAlignment = -4108     # xlCenter
                cell.WrapText = True
                sheet.Rows(at).RowHeight = high
                at += 1
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(at - 1, 2))
        for row in range(2, at):
            tops.append(round(sheet.Cells(row, 2).Top * 96 / 72))
            highs.append(round(sheet.Cells(row, 2).Height * 96 / 72))
        left = round(sheet.Cells(2, 2).Left * 96 / 72)
        wide = round(sheet.Cells(2, 2).Width * 96 / 72)
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(8):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.5)
            shot = ImageGrab.grabclipboard()
            if shot is not None:
                break
        else:
            print("Excel would not hand over a picture")
            return 1
        shot.save(SCRATCH / "excel.png")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()

    # The picture starts at the used range's top-left, so the offsets are
    # relative to it; both renders are read the same way.
    base = tops[0]
    truth = np.asarray(shot.convert("L"))
    theirs = offsets(truth, [t - base for t in tops], highs, 0, wide)

    # Oxi draws the whole used range, which starts further left and higher
    # than the rectangle Excel handed over, so its rows are taken from its own
    # dump rather than from Excel's geometry. Reading our picture at Excel's
    # coordinates is how this returned nothing but `None` the first time.
    out = SCRATCH / "oxi.png"
    subprocess.run([str(RENDERER), str(made), str(out)], capture_output=True, check=False)
    told = subprocess.run(
        [str(RENDERER), str(made), str(SCRATCH / "throwaway.png")],
        env={**os.environ, "OXI_XLSX_DUMP_ROWS": "1"},
        capture_output=True, text=True, encoding="utf-8",
    ).stdout
    ruled, down = {}, 0
    for line in told.splitlines():
        found = re.match(r"row (\d+) px (\d+)", line)
        if found:
            ruled[int(found.group(1))] = (down, int(found.group(2)))
            down += int(found.group(2))
    ours_img = np.asarray(Image.open(out).convert("L"))
    # Which column of pixels our column B occupies.
    our_left = 0
    for index in sorted(ruled):
        break
    told_columns = subprocess.run(
        [str(RENDERER), str(made), str(SCRATCH / "throwaway.png")],
        env={**os.environ, "OXI_XLSX_DUMP_COLUMNS": "1"},
        capture_output=True, text=True, encoding="utf-8",
    ).stdout
    across = 0
    for line in told_columns.splitlines():
        found = re.match(r"column (\d+) px (\d+)", line)
        if found:
            if int(found.group(1)) == 1:
                our_left = across
                break
            across += int(found.group(2))
    ours = []
    for at in range(2, 2 + len(tops)):
        top, high = ruled.get(at, (0, 0))
        block = ours_img[top:top + high, our_left:our_left + wide] < 140
        lit = np.where(block.any(axis=1))[0]
        ours.append(int(lit.min()) if len(lit) else None)

    print("  face          size letters  row px   Excel   Oxi   delta")
    at = 0
    seen = {}
    for face, size in ARMS:
      for words in WORDS:
        for _ in HIGHS:
            high = highs[at]
            a, b = theirs[at], ours[at]
            d = None if (a is None or b is None) else b - a
            seen[d] = seen.get(d, 0) + 1
            mark = "" if a == b else "   <--"
            print(f"  {face:<13}{size:>4.0f} {words:<8} {high:>6}   {str(a):>5}  {str(b):>4}"
                  f"   {('' if d is None else f'{d:+d}'):>5}{mark}")
            at += 1
    print("")
    print(f"  deltas seen: {seen}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
