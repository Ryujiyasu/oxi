# -*- coding: utf-8 -*-
r"""Does a cell's own border take room from its text?

The `1c*zbd` nine — the largest family still under 0.99 — put their figures
against the foot of rows whose bottom border is a DOUBLE, and every one of
them sits a pixel low in our picture. A double reaches one pixel INSIDE the
cell (`rule_for`), so the question is whether Excel lays its text in the box
the borders leave rather than in the whole cell.

One row an arm, the bottom border swept through every style Excel offers, the
text sat on the foot; a second column does the same against the top. The
reading is the ink's distance from the row's own edge, ours beside Excel's.

    python tools\metrics\_xlsx_border_room.py
    python tools\metrics\_xlsx_border_room.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_border_room")
FACE, POINTS = "ＭＳ Ｐゴシック", 11.0
COLUMN = 10.0
ROW_PT = 17.25          # `1c*zbd`'s own row height
WORDS = "0"
# (name, LineStyle, Weight) — the constants Excel's own object model uses.
STYLES = [
    ("none", -4142, None),
    ("hair", 1, 1),         # xlContinuous + xlHairline
    ("thin", 1, 2),         # xlContinuous + xlThin
    ("medium", 1, -4138),   # xlContinuous + xlMedium
    ("thick", 1, 4),        # xlContinuous + xlThick
    ("double", -4119, 4),   # xlDouble
    ("dotted", -4118, 2),   # xlDot
    ("dashed", -4115, 2),   # xlDash
]


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:F{len(STYLES) * 2 + 6}").Interior.Color = 0xFFFFFF
        for column in (2, 3, 4):
            sheet.Columns(column).ColumnWidth = COLUMN
        at = 2
        for _name, style, weight in STYLES:
            # bottom-aligned with a bottom rule, top-aligned with a top rule,
            # and centred with a bottom rule — the third says whether the rule
            # takes the pixel from the BOX or only from the foot.
            for column, edge, how in ((2, 9, -4107), (3, 8, -4160), (4, 9, -4108)):
                cell = sheet.Cells(at, column)
                cell.Value = 0
                cell.NumberFormat = "0"
                cell.Font.Name = FACE
                cell.Font.Size = POINTS
                cell.HorizontalAlignment = -4152        # right
                cell.VerticalAlignment = how            # bottom / top
                cell.Borders(edge).LineStyle = style
                if weight is not None and style != -4142:
                    cell.Borders(edge).Weight = weight
            sheet.Rows(at).RowHeight = ROW_PT
            # A blank row between arms, so one arm's border cannot be read as
            # the next one's edge.
            sheet.Rows(at + 1).RowHeight = ROW_PT
            at += 2
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(2, 2), sheet.Cells(at - 1, 4)).CopyPicture(
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
    made = SCRATCH / "borderroom.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    mine, columns, rows = ours(made)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    wide = columns[1][1] - columns[1][0]
    print(f"  {FACE} {POINTS}pt in a {round(ROW_PT * 96 / 72)}px row")
    print("  border  | on the foot: Excel Oxi | on the top: Excel Oxi | centred: Excel Oxi")
    walked = 0
    for at, (name, _style, _weight) in enumerate(STYLES):
        top, foot = rows[at * 2 + 2]
        tall = foot - top
        told = []
        for column in (0, 1, 2):
            # The reading is the DIGIT's own ink: the band skips two pixels at
            # each edge, where the border and the sheet's gridline are.
            band = truth[walked + 2:walked + tall - 2,
                         column * wide + 2:(column + 1) * wide - 2]
            lit = np.where(band.any(axis=1))[0]
            theirs = (int(lit.min()) + 2, int(lit.max()) + 2) if len(lit) else None
            left, right = columns[column + 1]
            band = mine[top + 2:foot - 2, left + 2:right - 2]
            lit = np.where(band.any(axis=1))[0]
            ours_at = (int(lit.min()) + 2, int(lit.max()) + 2) if len(lit) else None
            told.append((theirs, ours_at))
        walked += tall + (rows[at * 2 + 3][1] - rows[at * 2 + 3][0])
        if any(one is None for pair in told for one in pair):
            print(f"  {name:<8}|  nothing to read")
            continue
        print(f"  {name:<8}|"
              + "".join(f" {a[0]:>3}-{a[1]:<3} {b[0]:>3}-{b[1]:<3}{'' if a == b else '<<':<2}|"
                        for a, b in told))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
