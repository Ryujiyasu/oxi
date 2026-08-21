# -*- coding: utf-8 -*-
"""Where Excel puts a line of text inside its row, read at the gate's own scale.

The PDF answered this question with a scale factor of its own baked in, which
is fatal when the question is worth one pixel. This asks it the way the gate
asks everything else: Excel copies the sheet as a picture, Oxi draws the same
sheet, and the ink in each row band is compared. Rows vary only in font, size,
height and vertical alignment, so whatever moves is the rule.

    python tools\\metrics\\_xlsx_valign_shot.py            # build, shoot, draw, read
    python tools\\metrics\\_xlsx_valign_shot.py --reuse    # skip Excel, read again
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
BOOK = SCRATCH / "valign_shot.xlsx"
TRUTH = SCRATCH / "valign_shot.excel.png"
OURS = SCRATCH / "valign_shot.oxi.png"

FONTS = [("ＭＳ ゴシック", 10.0), ("ＭＳ ゴシック", 11.0), ("ＭＳ 明朝", 11.0),
         ("Calibri", 11.0), ("ＭＳ Ｐゴシック", 9.0), ("Meiryo UI", 12.0)]
# Points. None leaves the row at whatever the font makes it.
HEIGHTS = [None, 13.5, 18.0, 24.0, 36.0]
PLACES = ["top", "center", "bottom"]
TEXT = "Ag亜"


def build():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.title = "probe"
    sheet.column_dimensions["A"].width = 14
    probes, row = [], 1
    for face, points in FONTS:
        for height in HEIGHTS:
            for place in PLACES:
                cell = sheet.cell(row=row, column=1, value=TEXT)
                cell.font = Font(name=face, size=points)
                cell.alignment = Alignment(vertical=place, horizontal="left")
                if height is not None:
                    sheet.row_dimensions[row].height = height
                probes.append([row, face, points, place])
                row += 1
    book.save(BOOK)
    return probes


def heights_from_excel(probes):
    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(BOOK), 0, True)
        sheet = book.Worksheets(1)
        for probe in probes:
            probe.append(float(sheet.Rows(probe[0]).Height))
        book.Close(False)
    finally:
        excel.Quit()
    return probes


def shoot():
    listing = SCRATCH / "_batch.txt"
    listing.write_text(f"{BOOK.resolve()}\t{TRUTH.resolve()}", encoding="utf-8")
    TRUTH.unlink(missing_ok=True)
    subprocess.run(["powershell", "-NoProfile", "-File", str(SHOOTER),
                    "-ListFile", str(listing.resolve())],
                   capture_output=True, text=True, encoding="utf-8",
                   errors="replace", timeout=300)
    listing.unlink(missing_ok=True)


def draw():
    """Draws the sheet and returns the height Oxi gave each row.

    Its rows are asked for rather than assumed: where Oxi and Excel disagree
    about a height, bands built from one side's heights slide out of the
    other's, and every row below reads as misplaced text.
    """
    OURS.unlink(missing_ok=True)
    import os

    environment = dict(os.environ, OXI_XLSX_DUMP_ROWS="1")
    done = subprocess.run([str(RENDERER), str(BOOK), str(OURS), "96"],
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=300, env=environment)
    heights = {}
    for line in done.stdout.splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            # The dump counts rows the way the sheet does, from one.
            heights[int(parts[1])] = int(float(parts[3]))
    return heights


def ink_rows(image, top, bottom):
    """First and last row of the band that has any ink in it."""
    band = image[top:bottom]
    if band.size == 0:
        return None, None
    dark = np.flatnonzero((band < 200).any(axis=1))
    if dark.size == 0:
        return None, None
    return int(dark[0]) + top, int(dark[-1]) + top


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()

    probes = build()
    probes = heights_from_excel(probes)
    if not args.reuse:
        shoot()
    drawn = draw()

    truth = np.asarray(Image.open(TRUTH).convert("L"))
    ours = np.asarray(Image.open(OURS).convert("L"))
    print(f"Excel {truth.shape[1]}x{truth.shape[0]}, Oxi {ours.shape[1]}x{ours.shape[0]}")
    print()
    print(f"{'font':<15}{'pt':>4}{'row':>9}{'place':>8}"
          f"{'Excel ink':>12}{'Oxi ink':>12}{'top Δ':>7}{'bottom Δ':>9}")

    excel_at, our_at = 0, 0
    disagreed = {}
    for row, face, points, place, height in probes:
        pixels = int((height + 0.05) / 0.75)
        our_pixels = drawn.get(row, pixels)
        excel_top, excel_bottom = ink_rows(truth, excel_at, excel_at + pixels)
        our_top, our_bottom = ink_rows(ours, our_at, our_at + our_pixels)
        size = f"{pixels}" if our_pixels == pixels else f"{pixels}/{our_pixels}"
        if excel_top is None or our_top is None:
            print(f"{face:<15}{points:>4.0f}{size:>9}{place:>8}{'(no ink)':>12}")
        else:
            delta = (our_top - our_at) - (excel_top - excel_at)
            disagreed.setdefault((place, delta), []).append((face, points, pixels))
            print(f"{face:<15}{points:>4.0f}{size:>9}{place:>8}"
                  f"{excel_top - excel_at:>6}..{excel_bottom - excel_at:<5}"
                  f"{our_top - our_at:>6}..{our_bottom - our_at:<5}"
                  f"{delta:>+7}"
                  f"{(our_bottom - our_at) - (excel_bottom - excel_at):>+9}")
        excel_at += pixels
        our_at += our_pixels

    print()
    print("how the top of the ink differs, by placement:")
    for (place, delta) in sorted(disagreed):
        rows = disagreed[(place, delta)]
        print(f"   {place:<8}{delta:+d}px  ×{len(rows):<3} "
              f"{', '.join(f'{face} {points:.0f}/{pixels}px' for face, points, pixels in rows[:6])}"
              f"{' …' if len(rows) > 6 else ''}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
