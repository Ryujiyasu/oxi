# -*- coding: utf-8 -*-
r"""How far does a value axis stand off from the numbers beside it?

With the axis's own colour settled, what is left of the `zuhyo` family's
difference is where its numbers sit. Their ROWS agree exactly — 104..113,
148..157, and so on, on both sides — but Excel ends each number 18 pixels left
of the axis line where we end it 2. Four workbooks, every label, the same 16
pixels.

Sixteen pixels is a lot to guess at, and all four state the same 10pt label, so
the file cannot say whether the gap is a constant or grows with the text. This
asks the real chart: the same axis at six sizes, and at each of the three tick
styles, with the gap read straight off Excel's picture.

    python tools\metrics\_xlsx_chart_label_gap.py
    python tools\metrics\_xlsx_chart_label_gap.py --reuse
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import time
import zipfile
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
SOURCE = REPO / "tools" / "golden-test" / "documents" / "xlsx" / "a08feeb4a00b_zuhyo.xlsx"
SCRATCH = Path(r"C:\tmp\xlsx_chart_label_gap")
RENDERER = (
    REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
)
# What the pixel diff hands the renderer for this workbook.
RANGE = "1,1,53,87"
PART = "xl/charts/chart1.xml"
# Where the value axis stands in the untouched picture, and how far left of it
# to look for the numbers.
AXIS, REACH = 86, 70
# The rows the plot itself covers, so the axis is not confused with the sheet.
PLOT = (60, 520)


def sized(points: float):
    """Set the size of the value axis's own labels."""

    def alter(chart: str) -> str:
        found = re.search(r"<c:valAx>.*?</c:valAx>", chart, re.S)
        axis = found.group(0)
        said = re.sub(
            r'(<c:txPr>.*?<a:defRPr[^>]*?)sz="\d+"',
            lambda m: f'{m.group(1)}sz="{int(points * 100)}"',
            axis,
            count=1,
            flags=re.S,
        )
        return chart.replace(axis, said)

    return alter


def ticked(kind: str):
    """Set which way the axis's major tick marks point."""

    def alter(chart: str) -> str:
        found = re.search(r"<c:valAx>.*?</c:valAx>", chart, re.S)
        axis = found.group(0)
        said = re.sub(
            r'<c:majorTickMark val="[^"]*"/>',
            f'<c:majorTickMark val="{kind}"/>',
            axis,
            count=1,
        )
        return chart.replace(axis, said)

    return alter


ARMS: list[tuple[str, object]] = [
    ("as it stands (10pt, in)", lambda one: one),
    ("6pt", sized(6.0)),
    ("7pt", sized(7.0)),
    ("8pt", sized(8.0)),
    ("9pt", sized(9.0)),
    ("11pt", sized(11.0)),
    ("12pt", sized(12.0)),
    ("14pt", sized(14.0)),
    ("16pt", sized(16.0)),
    ("18pt", sized(18.0)),
    ("20pt", sized(20.0)),
    ("24pt", sized(24.0)),
    ("tick out", ticked("out")),
    ("tick none", ticked("none")),
    ("tick cross", ticked("cross")),
]


def build(made: Path, alter) -> None:
    if made.exists():
        made.unlink()
    with zipfile.ZipFile(SOURCE) as was, zipfile.ZipFile(made, "w", zipfile.ZIP_DEFLATED) as now:
        for item in was.infolist():
            held = was.read(item.filename)
            if item.filename == PART:
                held = alter(held.decode("utf-8")).encode("utf-8")
            now.writestr(item, held)


def shoot(made: Path, into: Path) -> bool:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(made))
    try:
        sheet = book.Worksheets(1)
        used = sheet.UsedRange
        for _ in range(8):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(1.2)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(into)
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def read(picture: Path) -> str:
    """Where the axis stands, and how far the numbers beside it stop short."""
    grey = np.asarray(Image.open(picture).convert("L")).astype(int)
    ink = grey < 128
    # The axis is the tallest run INSIDE the plot's own band, and it is found
    # at a LIGHTER threshold than the labels: this chart's axis states no
    # colour and Excel draws it 137,137,137, which is not ink by any test the
    # black numbers would pass. Bounding the band matters too — taken from the
    # whole column the tallest run is whatever else the sheet holds.
    plot = (grey < 200)[PLOT[0] : PLOT[1]]
    axis, best = AXIS, 0
    for x in range(40, min(ink.shape[1] // 2, 320)):
        run = int(plot[:, x].sum())
        if run > best:
            best, axis = run, x
    if best < 250:
        return f"no axis in the plot band (best run {best} at x={axis})"
    band = ink[PLOT[0] : PLOT[1], max(0, axis - REACH) : axis - 1]
    rows = band.sum(axis=1)
    lit = [y for y, v in enumerate(rows) if v > 3]
    runs, start, last = [], None, None
    for y in lit:
        if start is None:
            start = last = y
        elif y > last + 2:
            runs.append((start, last))
            start = last = y
        else:
            last = y
    if start is not None:
        runs.append((start, last))
    gaps = []
    for y0, y1 in runs:
        if y1 - y0 < 5:
            continue
        cols = np.nonzero(band[y0 : y1 + 1].any(axis=0))[0]
        if len(cols):
            gaps.append(axis - (max(0, axis - REACH) + int(cols[-1])))
    if not gaps:
        return f"axis x={axis}, no labels read"
    common = max(set(gaps), key=gaps.count)
    return f"gap {common:>3} of {sorted(set(gaps))}"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    SCRATCH.mkdir(parents=True, exist_ok=True)
    for at, (name, alter) in enumerate(ARMS):
        made = SCRATCH / f"arm{at}.xlsx"
        shot = SCRATCH / f"arm{at}.png"
        if not args.reuse:
            build(made, alter)
            if not shoot(made, shot):
                print(f"  {name:<24} Excel would not hand over a picture")
                continue
        if not shot.exists():
            print(f"  {name:<24} no picture")
            continue
        ours = SCRATCH / f"arm{at}.oxi.png"
        drawing = dict(os.environ)
        drawing["OXI_XLSX_RANGE"] = RANGE
        subprocess.run(
            [str(RENDERER), str(made), str(ours), "96"],
            capture_output=True, text=True, encoding="utf-8", env=drawing,
        )
        mine = read(ours) if ours.exists() else "not drawn"
        print(f"  {name:<24} E {read(shot):<34} O {mine}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
