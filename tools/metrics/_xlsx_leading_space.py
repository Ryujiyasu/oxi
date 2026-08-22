# -*- coding: utf-8 -*-
"""Does Excel draw a space at the start of a cell's text?

`h2daa2023_dendeba_kmc` holds `" ご覧になりたい…"` — a real space, kept by
`xml:space="preserve"` — and Excel starts the ink six pixels left of where the
renderer does, which is about what that space is worth. Either Excel drops it
or it is worth less than GDI says. This asks with the same string written both
ways, at three indents, so the answer does not depend on the indent as well.

    python tools\\metrics\\_xlsx_leading_space.py
    python tools\\metrics\\_xlsx_leading_space.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_leading_space")
BOOK = SCRATCH / "leading_space.xlsx"

FONTS = [("游ゴシック", 11.0), ("ＭＳ Ｐゴシック", 11.0)]
INDENTS = [0, 2]
TEXTS = [("plain", "あい"), ("one space", " あい"), ("two spaces", "  あい"),
         ("wide space", "\u3000あい"), ("trailing", "あい ")]


def build():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["A"].width = 30.0
    cases, row = [], 1
    for face, points in FONTS:
        for indent in INDENTS:
            for name, text in TEXTS:
                cell = sheet.cell(row=row, column=1, value=text)
                cell.font = Font(name=face, size=points)
                cell.alignment = Alignment(horizontal="left", vertical="center",
                                           indent=indent)
                sheet.row_dimensions[row].height = 18.0
                cases.append((row, face, points, indent, name))
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
    ours_png = SCRATCH / "leading_space.oxi.png"
    environment = dict(os.environ, OXI_XLSX_DUMP_ROWS="1")
    done = subprocess.run([str(RENDERER), str(BOOK), str(ours_png), "96"],
                          capture_output=True, timeout=300, env=environment)
    heights = {}
    for line in done.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            heights[int(parts[1])] = int(float(parts[3]))
    if not picture.exists():
        print("Excel gave no picture")
        return
    truth = np.asarray(Image.open(picture).convert("L"))
    ours = np.asarray(Image.open(ours_png).convert("L"))

    edges, at = {}, 0
    for index in sorted(heights):
        edges[index] = (at, at + heights[index])
        at += heights[index]

    print(f"{'font':<16}{'pt':>5}{'indent':>7}{'text':>12}"
          f"{'Excel ink':>12}{'ours':>8}{'差':>6}")
    for row, face, points, indent, name in cases:
        if row not in edges:
            continue
        top, foot = edges[row]
        if foot > min(truth.shape[0], ours.shape[0]):
            continue
        theirs = np.flatnonzero((truth[top:foot] < 128).sum(axis=0))
        mine = np.flatnonzero((ours[top:foot] < 128).sum(axis=0))
        if theirs.size == 0 or mine.size == 0:
            continue
        print(f"{face:<16}{points:>5}{indent:>7}{name:>12}"
              f"{int(theirs[0]):>12}{int(mine[0]):>8}"
              f"{int(mine[0]) - int(theirs[0]):>+6}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
