# -*- coding: utf-8 -*-
"""Where does Excel put a right-aligned line when the face is slanted?

`R6kessan` is 234 cells of ＭＳ Ｐゴシック 11pt bold ITALIC numbers against the
right edge, and every one of them sits two pixels right of Excel's. ＭＳ Ｐ
ゴシック has neither a bold nor an italic cut, so both are synthesised — GDI
does its own, Excel does its own, and the question is whether the difference is
in the ADVANCE (which moves a right-aligned line) or only in the ink.

One cell an arm, the four dressings, read against the cell's own right edge.

    python tools\\metrics\\_xlsx_italic_right.py
"""

from __future__ import annotations

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
SCRATCH = Path(r"C:\tmp\xlsx_italic_right")
WORDS = "0"   # one round glyph: its ink edges are the least ambiguous
# Each size twice: against the right edge, where the width Excel reserves
# shows, and against the left, where only the ink can differ.
ARMS = []
for _face in ("ＭＳ Ｐゴシック", "ＭＳ 明朝", "Century", "Times New Roman",
              "メイリオ", "游ゴシック", "Calibri", "Arial"):
    for _size in (11.0, 20.0):
        ARMS += [(_face, _size, False, True), (_face, _size, False, True)]
ROW_PT = 18.0
COLUMN = 40.0


def build(made: Path) -> int | None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:D{len(ARMS) + 4}").Interior.Color = 0xFFFFFF
        sheet.Columns(2).ColumnWidth = COLUMN
        for at, (face, points, bold, italic) in enumerate(ARMS, start=2):
            cell = sheet.Cells(at, 2)
            # A number, not text: a text number wears Excel's green
            # "stored as text" triangle, which is ink in the top-left corner
            # and would be read as the line's own start.
            cell.Value = 0
            cell.NumberFormat = "0"
            cell.Font.Name = face
            cell.Font.Size = points
            cell.Font.Bold = bold
            cell.Font.Italic = italic
            cell.HorizontalAlignment = -4152 if at % 2 == 0 else -4131  # right / left
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
            time.sleep(0.6)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                break
        else:
            return None
        return round(sheet.Cells(2, 2).Width * 96 / 72)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def ours(made: Path) -> tuple[np.ndarray, int, int, int]:
    import os
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
                          env={"OXI_XLSX_DUMP_COLUMNS": "1", "OXI_XLSX_DUMP_ROWS": "1",
                               **os.environ},
                          capture_output=True, text=True, encoding="utf-8")
    left, at, top, down = 0, 0, 0, 0
    for line in (told.stdout or "").splitlines() + (told.stderr or "").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column":
            if int(parts[1]) == 1:
                left = at
            at += int(parts[3])
        if len(parts) == 4 and parts[0] == "row":
            if int(parts[1]) == 2:
                top = down
            down += int(parts[3])
    return np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140, left, top, at


def main() -> int:
    made = SCRATCH / "italic.xlsx"
    wide = build(made)
    if wide is None:
        print("  Excel would not hand over a picture")
        return 1
    mine, left, top, _ = ours(made)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    tall = round(ROW_PT * 96 / 72)
    print(f"  {WORDS!r} against the right edge of a {wide}px column")
    print("  face             size  bold italic |  Excel ink   Oxi ink   start  end")
    for at, (face, points, bold, italic) in enumerate(ARMS):
        # The captured cell carries the sheet's own gridline at both edges, so
        # the reading skips two pixels either side and the band's top and foot.
        band = truth[at * tall + 2:(at + 1) * tall - 2, 1:wide - 1]
        lit = np.where(band.any(axis=0))[0]
        theirs = (int(lit.min()) + 1, int(lit.max()) + 1) if len(lit) else None
        band = mine[top + at * tall + 2:top + (at + 1) * tall - 2, left + 1:left + wide - 1]
        lit = np.where(band.any(axis=0))[0]
        ours_at = (int(lit.min()) + 1, int(lit.max()) + 1) if len(lit) else None
        if theirs is None or ours_at is None:
            print(f"  {face:<16}{points:>5}  {bold!s:>5} {italic!s:>6} |  nothing to read")
            continue
        print(f"  {face:<16}{points:>5}  {bold!s:>5} {italic!s:>6}"
              f" {'right' if at % 2 == 0 else 'left':<6}|"
              f"  {theirs[0]:>4}-{theirs[1]:<4} {ours_at[0]:>5}-{ours_at[1]:<4}"
              f"  {ours_at[0] - theirs[0]:>+5} {ours_at[1] - theirs[1]:>+4}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
