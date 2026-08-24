# -*- coding: utf-8 -*-
"""Where does Excel put a stacked character's ink, and where do we put ours?

`_xlsx_stack_face.py` reads the SHAPE of the ink: for all but nineteen marks
the upright face and the turned face put down the same glyph. This reads the
PLACE: the ink's own box inside the cell's, for Excel and for Oxi on the same
file, one character to a row.

Three cell shapes, because the corpus's headings are narrow enough that a
15-pixel em does not fit between the borders of a 17-pixel column:

    left     column 4.0, aligned left   — the pen's own offset, nothing else
    centre   column 4.0, aligned centre — the centring, with room to spare
    narrow   column 1.71, no alignment  — data_B01's own geometry

    python tools\\metrics\\_xlsx_stack_place.py
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
SCRATCH = Path(r"C:\tmp\xlsx_stack_place")

LETTERS = list("相談あウ一二ー～（）「」【】＝、。・：！？－／０１Ａ A1①ｱｰ")
ARMS = [("ＭＳ 明朝", 11.0), ("ＭＳ ゴシック", 8.0), ("ＭＳ 明朝", 14.0)]
SHAPES = {"left": (4.0, -4131), "centre": (4.0, -4108), "narrow": (1.71, None)}
ROW_PT = 30.0


def excel_shot(face: str, size: float, width: float, across: int | None,
               made: Path, shot: Path) -> int | None:
    """One character to a row, stacked, sat on the top of the row."""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:D{len(LETTERS) + 4}").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = width
        for at, letter in enumerate(LETTERS, start=2):
            cell = sheet.Cells(at, 2)
            cell.Value = letter
            cell.Font.Name = face
            cell.Font.Size = size
            cell.Orientation = -4166          # xlVertical
            cell.VerticalAlignment = -4160    # xlTop
            if across is not None:
                cell.HorizontalAlignment = across
            sheet.Rows(at).RowHeight = ROW_PT
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(LETTERS), 2))
        book.SaveAs(str(made), FileFormat=51)
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
        grabbed.save(shot)
        return round(sheet.Cells(2, 2).Width * 96 / 72)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def ruled(made: Path, what: str) -> dict[int, tuple[int, int]]:
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "x.png")],
                          env={**os.environ, what: "1"},
                          capture_output=True, text=True, encoding="utf-8").stdout
    kind = "column" if "COLUMN" in what else "row"
    out, at = {}, 0
    for line in told.splitlines():
        found = re.match(rf"{kind} (\d+) px (\d+)", line)
        if found:
            out[int(found.group(1))] = (at, at + int(found.group(2)))
            at += int(found.group(2))
    return out


def inked(block: np.ndarray) -> tuple[int, int, int, int] | None:
    rows, cols = np.where(block)
    if len(rows) == 0:
        return None
    return int(cols.min()), int(rows.min()), int(cols.max() - cols.min() + 1), int(rows.max() - rows.min() + 1)


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    for face, size in ARMS:
        for shape, (width, across) in SHAPES.items():
            made = SCRATCH / f"place_{shape}.xlsx"
            shot = SCRATCH / f"excel_{shape}.png"
            wide = excel_shot(face, size, width, across, made, shot)
            if wide is None:
                print("  Excel would not hand over a picture")
                return 1
            ours = SCRATCH / f"oxi_{shape}.png"
            subprocess.run([str(RENDERER), str(made), str(ours), "96"],
                           capture_output=True, check=False)
            rows = ruled(made, "OXI_XLSX_DUMP_ROWS")
            columns = ruled(made, "OXI_XLSX_DUMP_COLUMNS")
            truth = np.asarray(Image.open(shot).convert("L")) < 140
            drawn = np.asarray(Image.open(ours).convert("L")) < 140
            tall = round(ROW_PT * 96 / 72)
            left, right = columns.get(1, (0, 0))
            print(f"\n  {face} {size}pt — {shape}, column {wide}px "
                  f"(ours {right - left}px), row {tall}px")
            print("  char        Excel x,y w×h      Oxi x,y w×h      dx dy  dw dh")
            tally: dict[tuple[int, int], int] = {}
            for at, letter in enumerate(LETTERS):
                theirs = inked(truth[at * tall:(at + 1) * tall, :wide])
                top, bottom = rows.get(at + 2, (0, 0))
                mine = inked(drawn[top:bottom, left:right])
                shown = letter.encode("unicode_escape").decode("ascii")
                if theirs is None or mine is None:
                    print(f"   {shown:<10}  {str(theirs):>16}  {str(mine):>16}")
                    continue
                move = (mine[0] - theirs[0], mine[1] - theirs[1])
                tally[move] = tally.get(move, 0) + 1
                print(f"   {shown:<10}  {theirs[0]:>3},{theirs[1]:<3} {theirs[2]:>2}x{theirs[3]:<2}"
                      f"     {mine[0]:>3},{mine[1]:<3} {mine[2]:>2}x{mine[3]:<2}"
                      f"    {move[0]:+d} {move[1]:+d}  {mine[2] - theirs[2]:+d} {mine[3] - theirs[3]:+d}")
            print(f"   moves seen: {sorted(tally.items(), key=lambda kv: -kv[1])}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
