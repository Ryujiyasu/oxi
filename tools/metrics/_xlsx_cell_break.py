# -*- coding: utf-8 -*-
r"""Does a line break at the end of a cell's text spend a line?

`d79c5b0675b8_r03_seizosangyo_tkh` holds 「…（産業細分類別）」 with a break
after it, centred in a 54-pixel row. Excel puts that text where a single line
goes; this renderer counts two lines, so its block is twice as tall and the
words sit nine pixels high. The same shape of text is all over the corpus's
forms, so it is worth pinning: where does a break at the front, in the middle
and at the end leave the words, and what does the row grow to.

    python tools\metrics\_xlsx_cell_break.py
    python tools\metrics\_xlsx_cell_break.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_cell_break")
BOOK = SCRATCH / "break.xlsx"

FACE, POINTS = "ＭＳ Ｐゴシック", 11.0
TEXTS = [
    ("plain", "あA"),
    ("one trailing", "あA\n"),
    ("two trailing", "あA\n\n"),
    ("leading", "\nあA"),
    ("middle", "あA\nいB"),
    ("middle and trailing", "あA\nいB\n"),
]
# A row tall enough that where the block sits can be read, and one Excel is
# left to work out for itself.
HEIGHTS = [40.5, None]
PLACES = ["top", "center", "bottom"]
# `r03_seizosangyo_tkh`'s own cell does not wrap, and a cell that does not
# wrap may not honour the breaks it holds at all.
WRAPS = [True, False]


def build():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["A"].width = 24.0
    cases, row = [], 1
    for name, text in TEXTS:
        for wrap in WRAPS:
          for height in HEIGHTS:
            for place in PLACES:
                cell = sheet.cell(row=row, column=1, value=text)
                cell.font = Font(name=FACE, size=POINTS)
                # Wrapping is what makes Excel honour the breaks inside a
                # cell; without it the text is one line whatever it holds.
                cell.alignment = Alignment(vertical=place, horizontal="left",
                                           wrap_text=wrap)
                if height:
                    sheet.row_dimensions[row].height = height
                cases.append((f"{name}{'' if wrap else ' no wrap'}",
                              text, height, place, row))
                row += 1
    book.save(BOOK)
    return cases


def shoot():
    picture = BOOK.with_suffix(".excel.png")
    picture.unlink(missing_ok=True)
    listing = SCRATCH / "_batch.txt"
    listing.write_text(f"{BOOK.resolve()}\t{picture.resolve()}", encoding="utf-8-sig")
    subprocess.run(["powershell", "-NoProfile", "-File", str(SHOOTER),
                    "-ListFile", str(listing.resolve())],
                   capture_output=True, text=True, encoding="utf-8",
                   errors="replace", timeout=900)
    listing.unlink(missing_ok=True)
    return picture


def drawn():
    ours = SCRATCH / "break.oxi.png"
    done = subprocess.run(
        [str(RENDERER), str(BOOK), str(ours), "96"], capture_output=True, timeout=900,
        env=dict(os.environ, OXI_XLSX_DUMP_ROWS="1"))
    heights = {}
    for line in done.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            heights[int(parts[1])] = int(float(parts[3]))
    return ours, heights


def ink(picture, top, foot):
    band = (picture[top:foot] < 128)
    rows = np.flatnonzero(band.any(axis=1))
    return (int(rows[0]), int(rows[-1])) if rows.size else (None, None)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")

    cases = build()
    picture = BOOK.with_suffix(".excel.png") if args.reuse else shoot()
    if not picture.exists():
        print("Excel gave no picture")
        return
    truth = np.asarray(Image.open(picture).convert("L"))
    ours_png, heights = drawn()
    mine = np.asarray(Image.open(ours_png).convert("L"))
    edges, at = {}, 0
    for index in sorted(heights):
        edges[index] = (at, at + heights[index])
        at += heights[index]

    print(f"{'text':<20}{'height':>8}{'place':>8}{'row px':>8}"
          f"{'Excel ink':>12}{'ours':>10}")
    for name, _text, height, place, row in cases:
        if row not in edges:
            continue
        top, foot = edges[row]
        if foot > min(truth.shape[0], mine.shape[0]):
            continue
        theirs = ink(truth, top, foot)
        ours = ink(mine, top, foot)
        mark = "" if theirs == ours else "  <<"
        print(f"{name:<20}{'stated' if height else 'its own':>8}{place:>8}"
              f"{foot - top:>8}{f'{theirs[0]}..{theirs[1]}':>12}"
              f"{f'{ours[0]}..{ours[1]}':>10}{mark}")


if __name__ == "__main__":
    main()
