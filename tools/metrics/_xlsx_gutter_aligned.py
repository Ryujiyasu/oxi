# -*- coding: utf-8 -*-
"""How much room does Excel keep at the edge the text is aligned to?

The renderer keeps `3 + extra` before the text and `2 + extra` after it, with
`extra` from the cell font's own digit — derived from left-aligned text.
`h2dee1989kre`'s right-aligned numbers sit two pixels left of ours in every
column, which says the room at the right edge is not the same two pixels.

This puts the same string against the left, the middle and the right of a
column in several faces, and reads the gap either side off Excel's picture.

    python tools\\metrics\\_xlsx_gutter_aligned.py
    python tools\\metrics\\_xlsx_gutter_aligned.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_gutter_aligned")
BOOK = SCRATCH / "gutter.xlsx"

FACES = [("メイリオ", 11.0), ("游ゴシック", 11.0), ("ＭＳ Ｐゴシック", 11.0),
         ("ＭＳ ゴシック", 11.0), ("Meiryo UI", 11.0), ("Calibri", 11.0),
         ("メイリオ", 14.0), ("游ゴシック", 14.0), ("ＭＳ Ｐゴシック", 14.0),
         ("メイリオ", 8.0), ("游ゴシック", 8.0), ("ＭＳ Ｐゴシック", 20.0)]
PLACES = ["left", "center", "right"]
# A number, not a string: Excel marks a numeric-looking string with a
# green triangle in the corner of the cell, and that is ink too.
TEXT = 36044


def build():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["A"].width = 14.0
    cases, row = [], 1
    for face, points in FACES:
        for place in PLACES:
            cell = sheet.cell(row=row, column=1, value=TEXT)
            cell.font = Font(name=face, size=points)
            cell.alignment = Alignment(horizontal=place, vertical="center")
            sheet.row_dimensions[row].height = 21.0
            cases.append((row, face, points, place))
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


def ours():
    out = SCRATCH / "gutter.oxi.png"
    environment = dict(os.environ, OXI_XLSX_DUMP_ROWS="1", OXI_XLSX_DUMP_COLUMNS="1")
    done = subprocess.run([str(RENDERER), str(BOOK), str(out), "96"],
                          capture_output=True, timeout=300, env=environment)
    heights, width = {}, 0
    for line in done.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            heights[int(parts[1])] = int(float(parts[3]))
        if len(parts) == 4 and parts[0] == "column" and parts[1] == "0":
            width = int(float(parts[3]))
    return out, heights, width


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()

    cases = build()
    picture = BOOK.with_suffix(".excel.png") if args.reuse else shoot()
    ours_png, heights, width = ours()
    if not picture.exists():
        print("Excel gave no picture")
        return
    truth = np.asarray(Image.open(picture).convert("L"))
    mine = np.asarray(Image.open(ours_png).convert("L"))

    edges, at = {}, 0
    for index in sorted(heights):
        edges[index] = (at, at + heights[index])
        at += heights[index]

    print(f"the column is {width}px wide")
    print(f"{'face':<16}{'pt':>5}{'place':>8}"
          f"{'Excel left':>12}{'right':>7}{'ours left':>11}{'right':>7}")
    for row, face, points, place in cases:
        if row not in edges:
            continue
        top, foot = edges[row]
        if foot > min(truth.shape[0], mine.shape[0]):
            continue
        held = []
        for image in (truth, mine):
            band = image[top:foot, :width]
            ink = np.flatnonzero((band < 128).sum(axis=0))
            held.append((int(ink[0]), width - int(ink[-1]) - 1) if ink.size else (-1, -1))
        print(f"{face:<16}{points:>5}{place:>8}"
              f"{held[0][0]:>12}{held[0][1]:>7}{held[1][0]:>11}{held[1][1]:>7}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
