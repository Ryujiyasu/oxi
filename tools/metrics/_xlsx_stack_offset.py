"""The offset of a stacked letter inside its line box, per face and per size.

`_xlsx_stack_block.py` settled the block: `n * pitch`, face-independent, with
the spare pixel above. What is left is where the letter sits INSIDE its box,
and that is the piece no expression reproduces — Excel's effective em runs
11, 12, 13, 14, 16, 18 at 8/9/10/11/12/14pt while `size * 4 / 3` gives
10.67 … 18.67, so rounding is right at one end and truncation at the other.
`TEXTMETRIC` does not explain it either: GDI hands back the pixel size it was
asked for, identically for the `@` face and the plain one.

So it is measured rather than derived, the way `row_defaults` is. Sitting the
text on the TOP of the cell pins the block to the row's top edge, which leaves
the offset alone in the reading. Only the twenty (face, size) pairs that
actually stack anywhere in the corpus are worth a row in the table; they are
swept here together with the rest of each face's sizes for headroom.

Run: python tools/metrics/_xlsx_stack_offset.py
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
SCRATCH = Path(r"C:\tmp\xlsx_stack_offset")
# Every face the corpus stacks anything in, and every size any of them uses.
FACES = ["ＭＳ 明朝", "ＭＳ ゴシック"]
SIZES = [6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0, 9.5, 10.0, 10.5, 11.0, 11.5,
         12.0, 12.5, 13.0, 14.0, 15.0, 16.0, 18.0, 20.0]
WORDS = "就業"
ROW_PX = 90


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "offset.xlsx"
    plan = [(face, size) for face in FACES for size in SIZES]
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:D80").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = 6.0
        for at, (face, size) in enumerate(plan, start=2):
            cell = sheet.Cells(at, 2)
            cell.Value = WORDS
            cell.Font.Name = face
            cell.Font.Size = size
            cell.Orientation = -4166           # xlVertical
            cell.VerticalAlignment = -4160     # xlTop
            cell.WrapText = True
            sheet.Rows(at).RowHeight = round(ROW_PX * 72 / 96, 2)
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(plan), 2))
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(0.8)
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

    subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png")],
                   capture_output=True, check=False)
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
    ours_img = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L"))

    def read(img, top, at_left):
        block = img[top:top + ROW_PX, at_left:at_left + wide] < 140
        lit = np.where(block.any(axis=1))[0]
        if not len(lit):
            return None, None
        rows_lit = block.any(axis=1)
        runs, start = [], None
        for y, on in enumerate(rows_lit):
            if on and start is None:
                start = y
            elif not on and start is not None:
                runs.append((start, y))
                start = None
        if start is not None:
            runs.append((start, len(rows_lit)))
        pitch = runs[1][0] - runs[0][0] if len(runs) > 1 else None
        return int(lit.min()), pitch

    print(f"  top-aligned, row {ROW_PX}px, letters {WORDS}")
    print("  face               size   Excel off/pitch   Oxi off/pitch   delta")
    seen = {}
    for at, (face, size) in enumerate(plan, start=2):
        theirs, their_pitch = read(truth, (at - 2) * ROW_PX, 0)
        ours, our_pitch = read(ours_img, ruled.get(at, 0), left)
        gap = None if (theirs is None or ours is None) else ours - theirs
        seen[gap] = seen.get(gap, 0) + 1
        flag = "" if gap == 0 else "   <--"
        print(f"  {face:<18}{size:>5.1f}   {str(theirs):>5} / {str(their_pitch):<5}"
              f"     {str(ours):>5} / {str(our_pitch):<5}"
              f"   {('' if gap is None else f'{gap:+d}'):>5}{flag}")
    print("")
    print(f"  deltas seen: {seen}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
