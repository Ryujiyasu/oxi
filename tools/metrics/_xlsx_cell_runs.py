# -*- coding: utf-8 -*-
r"""How does Excel set a wrapped cell whose text is dressed in pieces?

58 of the 285 workbooks hold 817 of them, `9fd461bf494a_zuhyo` and
`f1b851d0a096_001290291` among the corpus's worst. The renderer draws a
dressed run in its own font only when the cell comes out on one line; a
wrapped one it draws whole in the cell's own font, so a 20-point title inside
a 14-point cell comes out 14.

This asks Excel where such a cell's lines fall: whether a bigger piece grows
the line it sits on, where the baseline of a mixed line is, and how many lines
the text takes when the pieces are measured in their own fonts.

    python tools\metrics\_xlsx_cell_runs.py
    python tools\metrics\_xlsx_cell_runs.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_cell_runs")
BOOK = SCRATCH / "runs.xlsx"

FACE = "ＭＳ Ｐゴシック"
BASE = 11.0
# Each case is a list of (text, size, bold, face); the cell's own font is
# ＭＳ Ｐゴシック 11, and a piece that states nothing wears it.
CASES = [
    ("plain", [("あいうえおかきくけこさしすせそ", None, False, None)]),
    ("small then big", [("あいうえお", 9.0, False, None),
                        ("かきくけこさしす", 20.0, False, None)]),
    ("big then small", [("あいうえお", 20.0, False, None),
                        ("かきくけこさしす", 9.0, False, None)]),
    ("big in the middle", [("あいう", None, False, None),
                           ("えおかき", 18.0, False, None),
                           ("くけこさし", None, False, None)]),
    ("bold piece", [("あいうえお", None, False, None),
                    ("かきくけこ", None, True, None),
                    ("さしすせそ", None, False, None)]),
    ("another face", [("あいうえお", None, False, None),
                      ("かきくけこ", None, False, "ＭＳ 明朝"),
                      ("さしすせそ", None, False, None)]),
    ("big piece last", [("あいうえおかきくけこ", None, False, None),
                        ("さしす", 20.0, False, None)]),
]
# A column that holds about eight of the cell's own characters, so every case
# wraps. Every row states its height: a row Excel works out for itself comes
# out taller than this renderer's, and then the two pictures no longer share a
# row grid — every band below the first difference is measured against the
# wrong row (the trap that made the first run of this probe unreadable).
WIDTH = 18.0
HEIGHTS = [90.0]


def build():
    from openpyxl import Workbook
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    from openpyxl.styles import Alignment, Font

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["A"].width = WIDTH
    cases, row = [], 1
    for name, pieces in CASES:
        for height in HEIGHTS:
            held = CellRichText([
                TextBlock(InlineFont(rFont=face or FACE, sz=size or BASE,
                                     b=bold, charset=128), text)
                for text, size, bold, face in pieces
            ])
            cell = sheet.cell(row=row, column=1, value=held)
            cell.font = Font(name=FACE, size=BASE)
            cell.alignment = Alignment(vertical="top", horizontal="left",
                                       wrap_text=True)
            if height:
                sheet.row_dimensions[row].height = height
            cases.append((name, height, row))
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
    ours = SCRATCH / "runs.oxi.png"
    done = subprocess.run(
        [str(RENDERER), str(BOOK), str(ours), "96"], capture_output=True, timeout=900,
        env=dict(os.environ, OXI_XLSX_DUMP_ROWS="1"))
    heights = {}
    for line in done.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            heights[int(parts[1])] = int(float(parts[3]))
    return ours, heights


def bands(picture, top, foot):
    """Each line's (first row, last row, ink width) inside the band."""
    held = (picture[top:foot] < 128)
    rows = held.sum(axis=1)
    lines, run = [], None
    for step, lit in enumerate(rows):
        if lit:
            run = [step, step] if run is None else [run[0], step]
        elif run is not None and step - run[1] > 1:
            lines.append(tuple(run))
            run = None
    if run is not None:
        lines.append(tuple(run))
    out = []
    for first, last in lines:
        columns = np.flatnonzero(held[first:last + 1].any(axis=0))
        out.append((first, last, int(columns[-1] - columns[0] + 1) if columns.size else 0))
    return out


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

    for name, height, row in cases:
        if row not in edges:
            continue
        top, foot = edges[row]
        if foot > min(truth.shape[0], mine.shape[0]):
            continue
        theirs = bands(truth, top, foot)
        ours = bands(mine, top, foot)
        mark = "" if theirs == ours else "  <<"
        print(f"{name:<20}{'stated' if height else 'its own':>8} row {foot - top:>3}px{mark}")
        print(f"    Excel {theirs}")
        print(f"    ours  {ours}")


if __name__ == "__main__":
    main()
