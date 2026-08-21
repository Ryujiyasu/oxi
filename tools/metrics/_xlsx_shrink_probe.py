# -*- coding: utf-8 -*-
"""What size does Excel draw a cell at when it is told to shrink it to fit?

`shrinkToFit="1"` is asked for by 85 of the corpus's 285 workbooks and the
renderer does not read it at all, so those cells are drawn at full size and
run on or clip where Excel has quietly made them smaller. This asks Excel what
size it settles on: text of growing length in a column of fixed width, and the
ink read back against the same text drawn at every candidate size.

    python tools\\metrics\\_xlsx_shrink_probe.py
    python tools\\metrics\\_xlsx_shrink_probe.py --reuse
"""
import argparse
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _xlsx_wrap_probe import ink_extent, run_width  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
SHOOTER = Path(__file__).resolve().parent / "_xlsx_screen_shot.ps1"
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SCRATCH = Path(r"C:\tmp\xlsx_shrink")
BOOK = SCRATCH / "shrink.xlsx"
TRUTH = SCRATCH / "shrink.excel.png"

FONTS = [("ＭＳ Ｐゴシック", 11.0), ("ＭＳ ゴシック", 11.0), ("Calibri", 11.0)]
WIDTH = 10.0            # characters
SAMPLE = "あいうえおかきくけこさしすせそ"
LATIN = "abcdefghijklmno"
ROW_PX = 30             # pinned, so the row never grows to take the text


def build():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["A"].width = WIDTH
    probes, row = [], 1
    for face, points in FONTS:
        sample = LATIN if face == "Calibri" else SAMPLE
        for length in range(1, len(sample) + 1):
            cell = sheet.cell(row=row, column=1, value=sample[:length])
            cell.font = Font(name=face, size=points)
            cell.alignment = Alignment(shrink_to_fit=True, vertical="bottom",
                                       horizontal="left")
            sheet.row_dimensions[row].height = ROW_PX * 0.75
            probes.append((row, face, points, sample[:length]))
            row += 1
    book.save(BOOK)
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


def column_px():
    import os

    environment = dict(os.environ, OXI_XLSX_DUMP_COLUMNS="1")
    done = subprocess.run([str(RENDERER), str(BOOK), str(SCRATCH / "shrink.oxi.png"), "96"],
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=300, env=environment)
    for line in done.stdout.splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column":
            return int(float(parts[3]))
    return 0


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()

    probes = build()
    if not args.reuse:
        shoot()
    column = column_px()
    truth = np.asarray(Image.open(TRUTH).convert("L"))
    print(f"column {column}px, usable {column - 5}px")
    print(f"{'face':<14}{'pt':>5}{'chars':>6}{'at 11pt':>9}{'ink':>6}"
          f"{'Excel drew':>12}{'to fit':>9}")
    for index, (_, face, points, text) in enumerate(probes):
        top = index * ROW_PX
        strip = truth[top:top + ROW_PX, :column]
        lit = np.flatnonzero((strip < 128).any(axis=0))
        if lit.size == 0:
            continue
        reach = int(lit[-1]) - 3 + 1
        full = run_width(face, points, text)
        # Which size draws ink of that width: the sizes Excel can pick from,
        # a quarter point apart, down from the size the cell asks for.
        best, best_gap = points, 10_000
        size = 1.0
        while size <= points + 0.01:
            gap = abs(ink_extent(face, size, text) - reach)
            if gap < best_gap:
                best, best_gap = size, gap
            size += 0.25
        wanted = points * (column - 5) / full if full else points
        print(f"{face:<14}{points:>5.1f}{len(text):>6}{full:>9}{reach:>6}"
              f"{best:>12.2f}{min(points, wanted):>9.2f}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
