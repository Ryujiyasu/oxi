# -*- coding: utf-8 -*-
r"""What colour is a chart axis that states no colour?

`a08feeb4a00b_zuhyo` is near the corpus floor and its value axis is drawn grey
— 137,137,137 down its whole length — where we draw it black. The axis states
no `<c:spPr>` at all, so the colour is a default, and ours is black. Its
CATEGORY axis, which does state one (`w="3175"`, black), comes out solid black
on both sides, so this is not about the width.

Grey 137 could be a colour or it could be ink: a line thinner than a pixel
drawn at partial coverage would read the same from one sample. So the real
chart is asked directly — give the axis a colour it cannot be mistaken for,
then give it black, and see which of the two readings survives.

    python tools\metrics\_xlsx_chart_axis.py
    python tools\metrics\_xlsx_chart_axis.py --reuse
"""

from __future__ import annotations

import argparse
import re
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
SCRATCH = Path(r"C:\tmp\xlsx_chart_axis")
PART = "xl/charts/chart1.xml"


def with_line(said: str | None):
    """Give the value axis a line of its own, or take the one it has away."""

    def alter(chart: str) -> str:
        found = re.search(r"<c:valAx>.*?</c:valAx>", chart, re.S)
        if not found:
            return chart
        axis = re.sub(r"<c:spPr>.*?</c:spPr>", "", found.group(0), flags=re.S)
        if said is not None:
            # The order the schema wants: the axis's own properties come after
            # its tick marks and before its text properties.
            axis = axis.replace(
                "<c:txPr>", f"<c:spPr><a:ln>{said}</a:ln></c:spPr><c:txPr>", 1
            )
        return chart.replace(found.group(0), axis)

    return alter


BLACK = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
ARMS: list[tuple[str, object]] = [
    ("as it stands", lambda one: one),
    ("valAx red", with_line('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')),
    ("valAx black", with_line(BLACK)),
    ("valAx 898989", with_line('<a:solidFill><a:srgbClr val="898989"/></a:solidFill>')),
    ("valAx D9D9D9", with_line('<a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill>')),
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
    """The colour of the tallest vertical run of ink in the chart's left half."""
    rgb = np.asarray(Image.open(picture).convert("RGB")).astype(int)
    grey = rgb.sum(axis=2) // 3
    best, where = 0, None
    for x in range(40, min(grey.shape[1] // 2, 400)):
        run = int((grey[:, x] < 250).sum())
        if run > best:
            best, where = run, x
    if where is None:
        return "no axis found"
    lit = np.nonzero(grey[:, where] < 250)[0]
    middle = lit[len(lit) // 2]
    return f"x={where:>4} run={best:>4}  {tuple(int(one) for one in rgb[middle, where])}"


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
                print(f"  {name:<22} Excel would not hand over a picture")
                continue
        if not shot.exists():
            print(f"  {name:<22} no picture")
            continue
        print(f"  {name:<22} {read(shot)}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
