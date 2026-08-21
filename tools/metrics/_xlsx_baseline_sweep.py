# -*- coding: utf-8 -*-
"""Where the baseline sits in the line box, for every font the table knows.

The box a line of text stands in is the row a sheet of that font alone would
have, which `row_defaults.rs` already carries. Where the baseline sits inside
that box is the missing number: the device's descent above the bottom of the
box for the ＭＳ faces, Calibri and Arial, and a pixel or three higher for the
faces with a large internal leading. No formula over the font's own metrics
reproduced it, so — as with the heights themselves — it is measured.

Method: one row per entry, pinned to a known height with a capital H at the
top of the cell. Excel draws the sheet, Oxi draws it, and the two inks are
compared. Both engines were shown to lay the same glyph down at the same size
(_xlsx_glyph_size.py), so the difference between the two ink tops is the
difference between the two baselines, whatever either rasterizer does at the
edges — which is what makes this safe to read a pixel at a time.

    python tools\\metrics\\_xlsx_baseline_sweep.py            # measure
    python tools\\metrics\\_xlsx_baseline_sweep.py --emit     # rewrite the table
"""
import argparse
import re
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _xlsx_font_metrics import metrics  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
TABLE = REPO / "tools" / "oxi-xlsx-renderer" / "src" / "row_defaults.rs"
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SHOOTER = Path(__file__).resolve().parent / "_xlsx_screen_shot.ps1"
SCRATCH = Path(r"C:\tmp\xlsx_baseline")
SLACK = 12          # pixels added to each row, so nothing is clipped or cramped
CHUNK = 60          # rows per workbook: Excel will not hand over a very tall picture


def entries():
    """The table as it stands: face, size in points, row height, baseline.

    The baseline is what the renderer is drawing with now — absent before this
    sweep first ran, and present afterwards, which is what makes a second run
    a check rather than a repeat.
    """
    held = []
    for line in TABLE.read_text(encoding="utf-8").splitlines():
        found = re.search(r'\("([^"]+)",\s*(\d+),\s*(\d+)(?:,\s*(\d+))?\)', line)
        if found:
            held.append((found.group(1), int(found.group(2)) / 4.0,
                         int(found.group(3)),
                         int(found.group(4)) if found.group(4) else None))
    return held


def build(chunk, number):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["A"].width = 10
    for row, (face, points, px, _) in enumerate(chunk, start=1):
        cell = sheet.cell(row=row, column=1, value="H")
        cell.font = Font(name=face, size=points)
        cell.alignment = Alignment(vertical="top", horizontal="left")
        sheet.row_dimensions[row].height = (px + SLACK) * 0.75
    path = SCRATCH / f"baseline_{number:02d}.xlsx"
    book.save(path)
    return path


def shoot(books):
    listing = SCRATCH / "_batch.txt"
    lines = []
    for book in books:
        picture = book.with_suffix(".excel.png")
        picture.unlink(missing_ok=True)
        lines.append(f"{book.resolve()}\t{picture.resolve()}")
    listing.write_text("\n".join(lines), encoding="utf-8")
    subprocess.run(["powershell", "-NoProfile", "-File", str(SHOOTER),
                    "-ListFile", str(listing.resolve())],
                   capture_output=True, text=True, encoding="utf-8",
                   errors="replace", timeout=120 * len(books))
    listing.unlink(missing_ok=True)


def draw(book):
    ours = book.with_suffix(".oxi.png")
    ours.unlink(missing_ok=True)
    subprocess.run([str(RENDERER), str(book), str(ours), "96"],
                   capture_output=True, text=True, encoding="utf-8",
                   errors="replace", timeout=300)
    return ours


def ink_top(image, top, bottom):
    band = image[top:bottom]
    dark = np.flatnonzero((band < 128).any(axis=1))
    return None if dark.size == 0 else int(dark[0])


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--emit", action="store_true",
                        help="write the measured baselines back into the table")
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()

    SCRATCH.mkdir(parents=True, exist_ok=True)
    table = entries()
    chunks = [table[at:at + CHUNK] for at in range(0, len(table), CHUNK)]
    books = [build(chunk, number) for number, chunk in enumerate(chunks)]
    if not args.reuse:
        shoot(books)

    measured, missing = {}, []
    for chunk, book in zip(chunks, books):
        picture = book.with_suffix(".excel.png")
        if not picture.exists():
            missing.extend(chunk)
            continue
        truth = np.asarray(Image.open(picture).convert("L"))
        ours = np.asarray(Image.open(draw(book)).convert("L"))
        at = 0
        for face, points, px, stated in chunk:
            band = px + SLACK
            _, tm = metrics(face, points)
            theirs = ink_top(truth, at, at + band)
            mine = ink_top(ours, at, at + band)
            at += band
            if theirs is None or mine is None:
                missing.append((face, points, px))
                continue
            # What the renderer drew with, and what Excel evidently drew with.
            drawn = stated if stated is not None else px - tm.tmDescent
            measured[(face, points)] = (drawn + theirs - mine, drawn, px)

    print(f"{len(measured)} of {len(table)} entries measured"
          f"{f', {len(missing)} without ink' if missing else ''}")
    spread = {}
    for _, (found, drawn, _) in sorted(measured.items()):
        spread[drawn - found] = spread.get(drawn - found, 0) + 1
    print("how far Excel's baseline sits above the one the renderer drew with:")
    for gap in sorted(spread):
        print(f"   {gap:+d}px  ×{spread[gap]}")
    for (face, points), (found, drawn, px) in sorted(measured.items()):
        if found != drawn:
            print(f"   {face:<22}{points:>6.1f}  box {px:>3}  Excel {found:>3}"
                  f"  drawn {drawn}")

    if args.emit:
        emit(table, measured)


def emit(table, measured):
    """Rewrite row_defaults.rs with the baseline beside each height."""
    text = TABLE.read_text(encoding="utf-8")
    lines, done = [], 0
    for line in text.splitlines():
        found = re.search(r'^(\s*)\("([^"]+)",\s*(\d+),\s*(\d+)(?:,\s*\d+)?\),\s*$', line)
        if not found:
            lines.append(line)
            continue
        face, quarters, px = found.group(2), int(found.group(3)), int(found.group(4))
        points = quarters / 4.0
        held = measured.get((face, points))
        baseline = held[0] if held else px - metrics(face, points)[1].tmDescent
        lines.append(f'{found.group(1)}("{face}", {quarters}, {px}, {baseline}),')
        done += 1
    TABLE.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"rewrote {done} entries in {TABLE}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
