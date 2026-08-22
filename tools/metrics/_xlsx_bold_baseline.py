# -*- coding: utf-8 -*-
"""Does a bold face sit where its regular does?

`row_defaults.rs` is keyed by face and size alone, so a bold cell is laid out
on the regular's line box and baseline. For most families that is right — the
bold is a synthesised thickening of the same outlines — but 游ゴシック,
メイリオ and their kin ship a separate bold design with metrics of its own,
and `mappingsheet2025`'s bold URLs come out a pixel high. This puts the same
letter on the sheet in both weights and reads the ink top off Excel's picture.

    python tools\\metrics\\_xlsx_bold_baseline.py
    python tools\\metrics\\_xlsx_bold_baseline.py --reuse
"""
import argparse
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image

REPO = Path(__file__).resolve().parents[2]
SHOOTER = Path(__file__).resolve().parent / "_xlsx_screen_shot.ps1"
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SCRATCH = Path(r"C:\tmp\xlsx_bold_baseline")
BOOK = SCRATCH / "bold_baseline.xlsx"

FACES = ["游ゴシック", "メイリオ", "Yu Gothic UI", "ＭＳ Ｐゴシック", "ＭＳ ゴシック",
         "ＭＳ 明朝", "Meiryo UI", "Calibri", "Arial", "Aptos Narrow"]
SIZES = [9.0, 11.0, 14.0]
HEIGHT = 40.0       # points, tall enough that nothing is cramped


def build():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["A"].width = 10
    cases, row = [], 1
    for face in FACES:
        for points in SIZES:
            for bold in (False, True):
                cell = sheet.cell(row=row, column=1, value="H")
                cell.font = Font(name=face, size=points, bold=bold)
                cell.alignment = Alignment(vertical="top", horizontal="left")
                sheet.row_dimensions[row].height = HEIGHT
                cases.append((row, face, points, bold))
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()

    cases = build()
    picture = BOOK.with_suffix(".excel.png") if args.reuse else shoot()
    ours = BOOK.with_suffix(".oxi.png")
    import os
    done = subprocess.run([str(RENDERER), str(BOOK), str(ours), "96"],
                          capture_output=True, timeout=300,
                          env=dict(os.environ, OXI_XLSX_DUMP_ROWS="1"))
    heights = {}
    for line in done.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            heights[int(parts[1])] = int(float(parts[3]))
    edges, at = {}, 0
    for index in sorted(heights):
        edges[index] = at
        at += heights[index]
    if not picture.exists():
        print("Excel gave no picture")
        return
    truth = np.asarray(Image.open(picture).convert("L"))
    mine = np.asarray(Image.open(ours).convert("L"))

    print(f"{'face':<16}{'pt':>5}{'weight':>9}"
          f"{'Excel ink top':>15}{'ours':>7}{'差':>5}   bold - regular")
    was = {}
    for row, face, points, bold in cases:
        top = edges.get(row)
        if top is None:
            continue
        foot = top + heights[row]
        if foot > min(truth.shape[0], mine.shape[0]):
            continue
        def ink(image):
            band = image[top:foot]
            dark = np.flatnonzero((band < 128).any(axis=1))
            return None if dark.size == 0 else int(dark[0])
        theirs, ours_top = ink(truth), ink(mine)
        if theirs is None or ours_top is None:
            continue
        step = ""
        if bold:
            step = f"{theirs - was.get((face, points), theirs):+d}"
        else:
            was[(face, points)] = theirs
        print(f"{face:<16}{points:>5}{'bold' if bold else 'regular':>9}"
              f"{theirs:>15}{ours_top:>7}{ours_top - theirs:>5}   {step}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
