# -*- coding: utf-8 -*-
r"""Does Excel soften a shape's outline, and how wide is one that says nothing?

Two questions the corpus floor turns on, asked of the same picture.

**Soft or hard.** Read a column across `glossary_05`'s flowchart where a
connector runs and Excel gives two rows of exactly 127 where we give one row of
0. Two rows at half ink is a one-pixel line whose centre fell on the boundary
between them and was ANTI-ALIASED; a hard pen cannot produce it. The chart work
already found Excel rules a grid line hard and draws a series curve soft, so
which side a shape's own outline falls on decides whether 2176 of the corpus's
shapes are drawn with the right edge. This walks one line down through a pixel
in eighths and reads the ink it leaves at each stop.

**How wide.** Nineteen of `glossary_05`'s twenty shape outlines state no width
at all: they carry `<a:ln>` with a colour and nothing else, and an `<a:lnRef
idx="2">` beside it. So the width is the theme's, and the theme's second line
style is 12700 — one point, a third of a pixel wider than the 9525 we fall back
to. Arms 8 onward state a width, or state none against each `lnRef`, so what
Excel does with each can be read off rather than assumed.

    python tools\metrics\_xlsx_shape_softness.py
    python tools\metrics\_xlsx_shape_softness.py --reuse
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
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
RENDERER = (
    REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
)
SCRATCH = Path(r"C:\tmp\xlsx_shape_softness")

EMU_PT = 12700
EMU_PX = 9525
# What one cell comes to. The workbook is one Excel made, so its defaults are
# Excel's own rather than the ones a hand-written package falls back to — 72
# and 25 here against 64 and 20 there. Both are read off the picture instead
# of stated, since a wrong guess reads every arm at the wrong row and every
# arm then says the line is missing.
COL_PX, ROW_PX = 72, 25
LONG = 8            # columns the line runs across
ROW_STEP = 3        # rows between one arm and the next

# Each arm: what it says its width is, which of the theme's line styles it
# points at, and how far down into its row it starts, in eighths of a pixel.
# The first eight walk one stated width through a whole pixel; the rest ask
# what an unstated width comes to.
# The last four wear an arrowhead. `glossary_05` draws its whole flowchart out
# of headed connectors and those are the ones Excel gives two grey rows to, so
# whether a head changes how the LINE is drawn is the question they ask.
ARMS = (
    [(9525, None, eighth, False) for eighth in range(8)]
    + [(w, None, 0, False) for w in (6350, 12700, 19050, 25400)]
    + [(None, idx, 0, False) for idx in (None, 1, 2, 3)]
    + [(9525, None, eighth, True) for eighth in (0, 2, 4, 6)]
)


def drawing_xml() -> str:
    shapes = []
    for at, (width, lnref, eighth, headed) in enumerate(ARMS):
        row = 1 + at * ROW_STEP
        said = f' w="{width}"' if width is not None else ""
        style = (
            f"<xdr:style><a:lnRef idx=\"{lnref}\">"
            f"<a:schemeClr val=\"accent1\"/></a:lnRef>"
            f"<a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef>"
            f"<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>"
            f"<a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef>"
            f"</xdr:style>"
            if lnref is not None
            else ""
        )
        shapes.append(
            f"<xdr:twoCellAnchor>"
            f"<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{row}</xdr:row>"
            f"<xdr:rowOff>{int(eighth * EMU_PX / 8)}</xdr:rowOff></xdr:from>"
            f"<xdr:to><xdr:col>{1 + LONG}</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{row}</xdr:row>"
            f"<xdr:rowOff>{int(eighth * EMU_PX / 8)}</xdr:rowOff></xdr:to>"
            f"<xdr:cxnSp macro=\"\"><xdr:nvCxnSpPr>"
            f"<xdr:cNvPr id=\"{at + 2}\" name=\"rule {at}\"/>"
            f"<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>"
            f"<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>"
            f"<a:prstGeom prst=\"line\"><a:avLst/></a:prstGeom>"
            f"<a:ln{said}><a:solidFill>"
            f"<a:srgbClr val=\"000000\"/></a:solidFill>"
            + ("<a:tailEnd type=\"arrow\" w=\"med\" len=\"med\"/>" if headed else "")
            + f"</a:ln>"
            f"</xdr:spPr></xdr:cxnSp>{style}<xdr:clientData/></xdr:twoCellAnchor>"
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/'
        'spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/'
        'drawingml/2006/main">' + "".join(shapes) + "</xdr:wsDr>"
    )


def zoomed(sheet_xml: bytes, zoom: int) -> bytes:
    """Put a zoom on the sheet's view, or take the one it has away.

    `glossary_05` is saved at 85%, and Excel's screen picture is what the gate
    compares against — so whether the zoom reaches the shapes is a question the
    corpus asks and no other arm here does.
    """
    held = sheet_xml.decode("utf-8")
    held = re.sub(r' zoomScale(Normal|PageLayoutView)?="\d+"', "", held)
    if zoom != 100:
        held = held.replace(
            "<sheetView ", f'<sheetView zoomScale="{zoom}" zoomScaleNormal="{zoom}" ', 1
        )
    return held.encode("utf-8")


def build(made: Path, zoom: int) -> None:
    """A workbook Excel wrote, with its drawing part swapped for the arms.

    Written from scratch it would have no theme, and half of what is being
    asked here is what a shape inherits FROM the theme. So Excel makes the
    package — theme, styles, relationships and a drawing part with one shape
    in it — and only the drawing is replaced.
    """
    SCRATCH.mkdir(parents=True, exist_ok=True)
    seed = SCRATCH / "seed.xlsx"
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Shapes.AddLine(10.0, 10.0, 100.0, 10.0)
        if seed.exists():
            seed.unlink()
        book.SaveAs(str(seed), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    if made.exists():
        made.unlink()
    with zipfile.ZipFile(seed) as was, zipfile.ZipFile(made, "w", zipfile.ZIP_DEFLATED) as now:
        for item in was.infolist():
            held = was.read(item.filename)
            if item.filename == "xl/drawings/drawing1.xml":
                held = drawing_xml().encode("utf-8")
            if item.filename == "xl/worksheets/sheet1.xml":
                held = zoomed(held, zoom)
            now.writestr(item, held)


def shoot(made: Path) -> bool:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(made))
    try:
        sheet = book.Worksheets(1)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(f"A1:M{4 + len(ARMS) * ROW_STEP}").CopyPicture(
                    Appearance=1, Format=2
                )
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.9)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def measure(grey: np.ndarray, columns: int, rows: int) -> None:
    """Take the cell's size from the picture rather than trusting a default."""
    global COL_PX, ROW_PX
    COL_PX = (grey.shape[1] - 1) // columns
    ROW_PX = (grey.shape[0] - 1) // rows


def profile(grey: np.ndarray, at: int) -> str:
    """The ink a rule leaves, as the grey of the rows it touches."""
    row = 1 + at * ROW_STEP
    middle = row * ROW_PX
    left, right = COL_PX + 8, COL_PX * (1 + LONG) - 8
    if right > grey.shape[1]:
        return "off the picture"
    band = grey[max(0, middle - 3) : middle + 4, left:right]
    rows = [255 - float(one.mean()) for one in band]
    weight = sum(rows)
    return " ".join(f"{one:5.1f}" for one in rows) + f"   ink {weight:6.1f}"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    parser.add_argument("--zoom", type=int, default=100,
                        help="the zoom to save the sheet's view at")
    args = parser.parse_args()
    made = SCRATCH / "soft.xlsx"
    if not args.reuse:
        build(made, args.zoom)
        if not shoot(made):
            print("  Excel would not hand over a picture")
            return 1
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")).astype(float)
    drawing = dict(os.environ)
    drawing["OXI_XLSX_RANGE"] = f"1,1,{4 + len(ARMS) * ROW_STEP},13"
    subprocess.run(
        [str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", env=drawing,
    )
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")).astype(float)
    rows = 4 + len(ARMS) * ROW_STEP
    measure(truth, 13, rows)
    print(f"  Excel {truth.shape[1]}x{truth.shape[0]}, Oxi {mine.shape[1]}x{mine.shape[0]}"
          f"; a cell is {COL_PX} x {ROW_PX}")
    print(f"  {'said':>6} {'ref':>4} {'8th':>4} {'arw':>4}   Excel: ink per row about the rule")
    print(f"  {'':>6} {'':>4} {'':>4}   Oxi")
    for at, (width, lnref, eighth, headed) in enumerate(ARMS):
        said = "-" if width is None else f"{width / EMU_PT:.2f}pt"
        ref = "-" if lnref is None else str(lnref)
        print(f"  {said:>6} {ref:>4} {eighth:>4} {'yes' if headed else '-':>4}"
              f"   {profile(truth, at)}")
        print(f"  {'':>6} {'':>4} {'':>4} {'':>4}   {profile(mine, at)}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
