# -*- coding: utf-8 -*-
"""How wide is one level of Excel's indent, and which alignments obey it?

`alignment indent="2"` appears in 71 of the 285 workbooks — 709 non-zero uses
— and the renderer ignores it, which is why `h2daa2023_dendeba_kmc`'s lines
sit 24px left of Excel's with exactly the right width. This asks Excel how far
an indent pushes the text, against the cell's own font and size, and which
alignments take any notice of it.

    python tools\\metrics\\_xlsx_indent.py
    python tools\\metrics\\_xlsx_indent.py --reuse
"""
import argparse
import os
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image

REPO = Path(__file__).resolve().parents[2]
SHOOTER = Path(__file__).resolve().parent / "_xlsx_screen_shot.ps1"
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SCRATCH = Path(r"C:\tmp\xlsx_indent")
BOOK = SCRATCH / "indent.xlsx"

FONTS = [("ＭＳ Ｐゴシック", 11.0), ("ＭＳ Ｐゴシック", 14.0), ("ＭＳ Ｐゴシック", 8.0),
         ("ＭＳ ゴシック", 11.0), ("Meiryo UI", 11.0), ("Calibri", 11.0),
         ("游ゴシック", 11.0)]
INDENTS = [0, 1, 2, 3, 6]
PLACES = ["left", "right", "center", "distributed"]


def build(normal=None):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    if normal:
        # Whether one level of indent is worth the same in a workbook whose
        # Normal font is not the default.
        face, points = normal
        book._named_styles["Normal"].font = Font(name=face, size=points)
    sheet = book.active
    sheet.column_dimensions["A"].width = 30.0
    cases, row = [], 1
    for face, points in FONTS:
        for indent in INDENTS:
            for place in PLACES:
                cell = sheet.cell(row=row, column=1, value="あ")
                cell.font = Font(name=face, size=points)
                cell.alignment = Alignment(horizontal=place, vertical="center",
                                           indent=indent)
                sheet.row_dimensions[row].height = 18.0
                cases.append((row, face, points, indent, place))
                row += 1
    book.save(BOOK)
    return cases


def shoot():
    picture = BOOK.with_suffix(".excel.png")
    picture.unlink(missing_ok=True)
    listing = SCRATCH / "_batch.txt"
    listing.write_text(f"{BOOK.resolve()}\t{picture.resolve()}", encoding="utf-8")
    subprocess.run(["powershell", "-NoProfile", "-File", str(SHOOTER),
                    "-ListFile", str(listing.resolve())],
                   capture_output=True, text=True, encoding="utf-8",
                   errors="replace", timeout=600)
    listing.unlink(missing_ok=True)
    return picture


def rows_of():
    ours = SCRATCH / "indent.oxi.png"
    environment = dict(os.environ, OXI_XLSX_DUMP_ROWS="1", OXI_XLSX_DUMP_COLUMNS="1")
    done = subprocess.run([str(RENDERER), str(BOOK), str(ours), "96"],
                          capture_output=True, timeout=300, env=environment)
    heights, width = {}, None
    for line in done.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            heights[int(parts[1])] = int(float(parts[3]))
        if len(parts) == 4 and parts[0] == "column" and parts[1] == "0":
            width = int(float(parts[3]))
    return ours, heights, width


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    parser.add_argument("--normal", help="the workbook's Normal font, as face:points")
    args = parser.parse_args()

    normal = None
    if args.normal:
        face, points = args.normal.rsplit(":", 1)
        normal = (face, float(points))
    cases = build(normal)
    picture = BOOK.with_suffix(".excel.png") if args.reuse else shoot()
    ours_png, heights, column = rows_of()
    if not picture.exists():
        print("Excel gave no picture")
        return
    truth = np.asarray(Image.open(picture).convert("L"))
    ours = np.asarray(Image.open(ours_png).convert("L"))

    edges, at = {}, 0
    for index in sorted(heights):
        edges[index] = (at, at + heights[index])
        at += heights[index]

    print(f"the column is {column}px wide")
    print(f"{'font':<16}{'pt':>5}{'indent':>7}{'place':>13}"
          f"{'Excel ink':>18}{'ours':>16}{'step':>7}")
    was = {}
    for row, face, points, indent, place in cases:
        if row not in edges:
            continue
        top, foot = edges[row]
        if foot > min(truth.shape[0], ours.shape[0]):
            continue
        theirs = np.flatnonzero((truth[top:foot] < 128).sum(axis=0))
        mine = np.flatnonzero((ours[top:foot] < 128).sum(axis=0))
        if theirs.size == 0 or mine.size == 0:
            continue
        # How far this indent moved the ink from the same case with none.
        edge = int(theirs[-1]) if place == "right" else int(theirs[0])
        step = "" if indent == 0 else f"{edge - was.get((face, points, place), edge):+d}"
        if indent == 0:
            was[(face, points, place)] = edge
        print(f"{face:<16}{points:>5}{indent:>7}{place:>13}"
              f"{f'{theirs[0]}..{theirs[-1]}':>18}{f'{mine[0]}..{mine[-1]}':>16}{step:>7}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
