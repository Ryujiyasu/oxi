# -*- coding: utf-8 -*-
"""Does an indent take room away from a wrapped cell's lines?

One level of indent moves the text 15px (`_xlsx_indent.py`). Whether it also
narrows the box the lines are broken in decides whether the indent belongs in
the row-height model as well as in the drawing. Excel is asked directly: the
same text in the same column at four indents, each cell ruled all the way
round and left to find its own height, so the rules in Excel's picture say how
many lines it took.

    python tools\\metrics\\_xlsx_indent_wrap.py
"""
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image

REPO = Path(__file__).resolve().parents[2]
SHOOTER = Path(__file__).resolve().parent / "_xlsx_screen_shot.ps1"
SCRATCH = Path(r"C:\tmp\xlsx_indent")
BOOK = SCRATCH / "indent_wrap.xlsx"

TEXT = "あいうえおかきくけこさしすせそたちつてと"
INDENTS = [0, 1, 2, 3]


def build(place="left"):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, Side

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["A"].width = 12.0
    edge = Side(style="thin", color="FF000000")
    frame = Border(left=edge, right=edge, top=edge, bottom=edge)
    for row, indent in enumerate(INDENTS, start=1):
        cell = sheet.cell(row=row, column=1, value=TEXT)
        cell.font = Font(name="ＭＳ Ｐゴシック", size=11)
        cell.alignment = Alignment(horizontal=place, vertical="top",
                                   wrap_text=True, indent=indent)
        cell.border = frame
    book.save(BOOK)


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
    place = sys.argv[1] if len(sys.argv) > 1 else "left"
    print(f"alignment {place}")
    build(place)
    picture = shoot()
    if not picture.exists():
        print("Excel gave no picture")
        return
    truth = np.asarray(Image.open(picture).convert("L"))
    height, width = truth.shape
    rules = [y for y in range(height) if (truth[y] < 128).sum() > 0.8 * width]
    # Neighbouring scanlines belong to the same rule.
    edges = [y for index, y in enumerate(rules) if index == 0 or y - rules[index - 1] > 1]
    print(f"picture {width}x{height}, rules at {edges}")
    for index, indent in enumerate(INDENTS):
        if index + 1 >= len(edges):
            break
        tall = edges[index + 1] - edges[index]
        print(f"  indent {indent}: {tall}px tall, about {tall / 18:.1f} lines")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
