"""Where does a stacked letter's ink sit inside its line box?

Centring mixes two unknowns: where the block starts, and where the ink sits
inside the block. Sitting the text on the TOP of the cell removes the first,
because the block starts at the row's top edge. What is left is the ink's own
offset, for Excel and for Oxi on the same file.

The pitch between stacked letters already agrees in every arm measured, so any
difference read here is the offset alone.

Run: python tools/metrics/_xlsx_stack_top.py
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
SCRATCH = Path(r"C:\tmp\xlsx_stack_top")
ARMS = [("ＭＳ 明朝", 8.0), ("ＭＳ ゴシック", 8.0), ("ＭＳ 明朝", 9.0),
        ("ＭＳ ゴシック", 10.0), ("ＭＳ Ｐゴシック", 11.0), ("ＭＳ 明朝", 11.0),
        ("ＭＳ ゴシック", 12.0), ("ＭＳ 明朝", 14.0)]
WORDS = "就業期間中"
ROW_PX = 120


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "top.xlsx"
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:D40").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = 5.0
        for at, (face, size) in enumerate(ARMS, start=2):
            cell = sheet.Cells(at, 2)
            cell.Value = WORDS
            cell.Font.Name = face
            cell.Font.Size = size
            cell.Orientation = -4166
            cell.VerticalAlignment = -4160     # xlTop
            cell.WrapText = True
            sheet.Rows(at).RowHeight = round(ROW_PX * 72 / 96, 2)
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(ARMS), 2))
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(0.6)
            shot = ImageGrab.grabclipboard()
            if shot is not None:
                break
        else:
            print("Excel would not hand over a picture")
            return 1
        shot.save(SCRATCH / "excel.png")
        wide = round(sheet.Cells(2, 2).Width * 96 / 72)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()

    out = SCRATCH / "oxi.png"
    subprocess.run([str(RENDERER), str(made), str(out)], capture_output=True, check=False)
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "x.png")],
                          env={**os.environ, "OXI_XLSX_DUMP_ROWS": "1"},
                          capture_output=True, text=True, encoding="utf-8").stdout
    ruled, down = {}, 0
    for line in told.splitlines():
        found = re.match(r"row (\d+) px (\d+)", line)
        if found:
            ruled[int(found.group(1))] = down
            down += int(found.group(2))
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "x.png")],
                          env={**os.environ, "OXI_XLSX_DUMP_COLUMNS": "1"},
                          capture_output=True, text=True, encoding="utf-8").stdout
    left, across = 0, 0
    for line in told.splitlines():
        found = re.match(r"column (\d+) px (\d+)", line)
        if found:
            if int(found.group(1)) == 1:
                left = across
                break
            across += int(found.group(2))

    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L"))
    ours_img = np.asarray(Image.open(out).convert("L"))

    def first(img, top, at_left):
        block = img[top:top + ROW_PX, at_left:at_left + wide] < 140
        lit = np.where(block.any(axis=1))[0]
        return int(lit.min()) if len(lit) else None

    print(f"  top-aligned, row {ROW_PX}px, letters {WORDS}")
    print("  face          size   Excel   Oxi   delta")
    seen = {}
    for at, (face, size) in enumerate(ARMS, start=2):
        theirs = first(truth, (at - 2) * ROW_PX, 0)
        ours = first(ours_img, ruled.get(at, 0), left)
        gap = None if (theirs is None or ours is None) else ours - theirs
        seen[gap] = seen.get(gap, 0) + 1
        print(f"  {face:<13}{size:>5.1f}  {str(theirs):>5}  {str(ours):>5}"
              f"   {('' if gap is None else f'{gap:+d}'):>5}")
    print("")
    print(f"  deltas seen: {seen}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
