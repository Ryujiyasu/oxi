# -*- coding: utf-8 -*-
"""How tall does one glyph come out, in Excel and in Oxi, at each point size?

Placement was measured first and the text turned out to be the wrong size:
ＭＳ ゴシック at ten point stands twelve pixels in Excel and fourteen here. A
Japanese face carries bitmaps for the small sizes and the two engines are
evidently picking different ones. One character per row, one row per size, and
the ink each of them leaves says which.

    python tools\\metrics\\_xlsx_glyph_size.py
"""
import argparse
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image

REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SHOOTER = Path(__file__).resolve().parent / "_xlsx_screen_shot.ps1"
SCRATCH = Path(r"C:\tmp\xlsx_valign")
BOOK = SCRATCH / "glyph_size.xlsx"
TRUTH = SCRATCH / "glyph_size.excel.png"
OURS = SCRATCH / "glyph_size.oxi.png"

FACES = ["ＭＳ ゴシック", "ＭＳ 明朝", "ＭＳ Ｐゴシック", "Meiryo UI", "Yu Gothic", "Calibri"]
SIZES = [8.0, 9.0, 10.0, 10.5, 11.0, 12.0, 14.0, 16.0, 18.0]
TEXT = "亜"          # fills the em square top to bottom in every Japanese face
LATIN = "H"          # a flat-topped capital, for the faces without kanji


def build():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["A"].width = 10
    probes, row = [], 1
    for face in FACES:
        for points in SIZES:
            letter = LATIN if face == "Calibri" else TEXT
            cell = sheet.cell(row=row, column=1, value=letter)
            cell.font = Font(name=face, size=points)
            cell.alignment = Alignment(vertical="top", horizontal="left")
            # Room enough that nothing is ever clipped or shifted by the box.
            sheet.row_dimensions[row].height = 30.0
            probes.append((row, face, points, letter))
            row += 1
    book.save(BOOK)
    return probes


def shoot():
    listing = SCRATCH / "_batch2.txt"
    listing.write_text(f"{BOOK.resolve()}\t{TRUTH.resolve()}", encoding="utf-8")
    TRUTH.unlink(missing_ok=True)
    subprocess.run(["powershell", "-NoProfile", "-File", str(SHOOTER),
                    "-ListFile", str(listing.resolve())],
                   capture_output=True, text=True, encoding="utf-8",
                   errors="replace", timeout=300)
    listing.unlink(missing_ok=True)


def draw():
    OURS.unlink(missing_ok=True)
    subprocess.run([str(RENDERER), str(BOOK), str(OURS), "96"],
                   capture_output=True, text=True, encoding="utf-8",
                   errors="replace", timeout=300)


def ink(image, top, bottom):
    band = image[top:bottom]
    rows = np.flatnonzero((band < 128).any(axis=1))
    columns = np.flatnonzero((band < 128).any(axis=0))
    if rows.size == 0:
        return None
    return int(rows[-1] - rows[0] + 1), int(columns[-1] - columns[0] + 1), int(rows[0])


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()

    probes = build()
    if not args.reuse:
        shoot()
    draw()
    truth = np.asarray(Image.open(TRUTH).convert("L"))
    ours = np.asarray(Image.open(OURS).convert("L"))

    print(f"{'face':<15}{'pt':>6}{'px asked':>10}"
          f"{'Excel h×w':>12}{'Oxi h×w':>10}{'Δh':>5}{'Δw':>4}{'Δtop':>6}")
    band = 40                       # 30pt rows
    for index, (_, face, points, letter) in enumerate(probes):
        top = index * band
        left = ink(truth, top, top + band)
        right = ink(ours, top, top + band)
        asked = points * 96.0 / 72.0
        if left is None or right is None:
            print(f"{face:<15}{points:>6.1f}{asked:>10.2f}{'(no ink)':>12}")
            continue
        print(f"{face:<15}{points:>6.1f}{asked:>10.2f}"
              f"{left[0]:>7}×{left[1]:<4}{right[0]:>5}×{right[1]:<4}"
              f"{right[0] - left[0]:>+5}{right[1] - left[1]:>+4}{right[2] - left[2]:>+6}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
