# -*- coding: utf-8 -*-
"""What does a stacked cell centre its characters on?

The turned pen sits the face's descent to the right of the standing one — read
inside ONE cell by `_xlsx_stack_turnpen.py`, so the cell's own centring cancels.
A cell holding a turned mark ALONE lands a pixel further left than that rule
predicts, which says the centring itself depends on what is in the cell: the
turned face's box is two pixels wider than the em at 11 point.

So: four cells, one character order each, and where every character lands.

    相相   both standing
    相（   standing first
    （相   turned first
    （）   both turned

    python tools\\metrics\\_xlsx_stack_centre.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _xlsx_stack_pen import offset  # noqa: E402
from _xlsx_stack_turnpen import metrics  # noqa: E402

Image.MAX_IMAGE_PIXELS = None
SCRATCH = Path(r"C:\tmp\xlsx_stack_centre")
TA_TOP, TA_BASELINE = 0, 24
STACKS = ["相相", "相（", "（相", "（）"]
FACE, SIZE = "ＭＳ 明朝", 11.0
ROW_PT = 45.0
COLUMN = 6.0


def shot() -> tuple[np.ndarray, int, int] | None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:H8").Interior.Color = 0xFFFFFF
        for at in range(2, 2 + len(STACKS)):
            sheet.Columns(at).ColumnWidth = COLUMN
            sheet.Columns(at).NumberFormat = "@"
        for column, held in enumerate(STACKS, start=2):
            cell = sheet.Cells(2, column)
            cell.Value = held
            cell.Font.Name = FACE
            cell.Font.Size = SIZE
            cell.Orientation = -4166          # xlVertical
            cell.VerticalAlignment = -4160    # xlTop
            cell.HorizontalAlignment = -4108  # xlCenter
        sheet.Rows(2).RowHeight = ROW_PT
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(2, 1 + len(STACKS)))
        book.SaveAs(str(SCRATCH / "centre.xlsx"), FileFormat=51)
        for _ in range(12):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(0.6)
            grabbed = ImageGrab.grabclipboard()
            if grabbed is not None:
                break
        else:
            return None
        grabbed.save(SCRATCH / "excel.png")
        wide = round(sheet.Cells(2, 2).Width * 96 / 72)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    return truth, wide, round(ROW_PT * 96 / 72)


def ink(block: np.ndarray) -> tuple[int, int] | None:
    rows, cols = np.where(block)
    if len(rows) == 0:
        return None
    return int(cols.min()), int(rows.min())


def main() -> int:
    got = shot()
    if got is None:
        print("  Excel would not hand over a picture")
        return 1
    truth, wide, tall = got
    px = round(SIZE * 96 / 72)
    ascU, descU, _ = metrics(FACE, px, False)
    _, _, tmH = metrics(FACE, px, True)
    on_em = round((wide - px) / 2)
    on_turned = round((wide - tmH) / 2)
    print(f"  {FACE} {SIZE}pt — {px}px em, turned box {tmH}px, column {wide}px")
    print(f"  centred on the em: {on_em}   centred on the turned box: {on_turned}"
          f"   descent {descU}")
    print("  stack   character  ink x   pen x   pen - on_em")
    for column, held in enumerate(STACKS):
        band = truth[:tall, column * wide:(column + 1) * wide]
        head = ink(band)
        if head is None:
            print(f"  {held}  nothing to read")
            continue
        guide = offset(FACE, px, held[0], held[0] == "（", TA_BASELINE if held[0] == "（" else TA_TOP)
        below = ink(band[head[1] + 13:])
        for which, (letter, seen) in enumerate((
            (held[0], head),
            (held[1], None if below is None else (below[0], below[1] + head[1] + 13)),
        )):
            if seen is None:
                continue
            turned = letter == "（"
            off = offset(FACE, px, letter, turned, TA_BASELINE if turned else TA_TOP)
            pen = seen[0] - off[0]
            print(f"  {held}    {letter}       {seen[0]:>5}   {pen:>5}   {pen - on_em:>+11d}")
        _ = guide
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
