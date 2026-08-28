# -*- coding: utf-8 -*-
r"""A line inside a stretched group: does Excel put it on a whole pixel?

`_xlsx_shape_softness.py` established that a shape's outline standing on its
own is HARD — one solid row of ink, snapped to a whole pixel, at every eighth
of a pixel it was asked to stand at. Yet `glossary_05` has connectors Excel
draws as TWO rows of exactly 127, which is one pixel of ink spread over a
boundary, and that file holds four groups whose `chExt` is not their `ext`: the
children are stretched by about 1.15 vertically.

So the question is whether the stretch is what puts a line between pixels. Each
arm here is a group holding one horizontal line, stretched so that the line
lands a stated fraction of a pixel down. If Excel snaps, every arm comes back
as one solid row and the group is not the cause; if it softens, the fractions
show as pairs of grey rows and the cause is found.

    python tools\metrics\_xlsx_group_scale.py
    python tools\metrics\_xlsx_group_scale.py --reuse
"""

from __future__ import annotations

import argparse
import os
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
SCRATCH = Path(r"C:\tmp\xlsx_group_scale")

EMU_PX = 9525
LONG = 8            # columns the line runs across
ROW_STEP = 4        # rows between one arm and the next
COL_PX, ROW_PX = 72, 25
# How far down a pixel each arm's line is meant to land, in eighths. The group
# carries it there: the child is written on a whole pixel and the stretch is
# what moves it.
ARMS = list(range(8))
SCALE = 2           # the group is twice its child space, top to bottom


def drawing_xml() -> str:
    shapes = []
    for at, eighth in enumerate(ARMS):
        row = 1 + at * ROW_STEP
        # The group covers two rows of its own; its children live in a space
        # half as tall, so everything in it is drawn at twice the distance from
        # the group's top. A child a quarter pixel down comes out half a pixel
        # down, which is how the fractions are reached without writing one.
        top = row * ROW_PX * EMU_PX
        child_top = 100 * EMU_PX
        child_tall = ROW_PX * EMU_PX
        down = int(eighth * EMU_PX / (8 * SCALE))
        shapes.append(
            f"<xdr:twoCellAnchor>"
            f"<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            f"<xdr:to><xdr:col>{1 + LONG}</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{row + 2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
            f"<xdr:grpSp><xdr:nvGrpSpPr>"
            f"<xdr:cNvPr id=\"{100 + at}\" name=\"group {at}\"/>"
            f"<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm>"
            f"<a:off x=\"0\" y=\"{top}\"/>"
            f"<a:ext cx=\"{LONG * COL_PX * EMU_PX}\" cy=\"{child_tall * SCALE}\"/>"
            f"<a:chOff x=\"0\" y=\"{child_top}\"/>"
            f"<a:chExt cx=\"{LONG * COL_PX * EMU_PX}\" cy=\"{child_tall}\"/>"
            f"</a:xfrm></xdr:grpSpPr>"
            f"<xdr:cxnSp macro=\"\"><xdr:nvCxnSpPr>"
            f"<xdr:cNvPr id=\"{200 + at}\" name=\"rule {at}\"/>"
            f"<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>"
            f"<a:xfrm><a:off x=\"0\" y=\"{child_top + down}\"/>"
            f"<a:ext cx=\"{LONG * COL_PX * EMU_PX}\" cy=\"0\"/></a:xfrm>"
            f"<a:prstGeom prst=\"line\"><a:avLst/></a:prstGeom>"
            f"<a:ln w=\"9525\"><a:solidFill>"
            f"<a:srgbClr val=\"000000\"/></a:solidFill></a:ln>"
            f"</xdr:spPr></xdr:cxnSp></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor>"
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/'
        'spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/'
        'drawingml/2006/main">' + "".join(shapes) + "</xdr:wsDr>"
    )


def build(made: Path) -> None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    seed = SCRATCH / "seed.xlsx"
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        book.Worksheets(1).Shapes.AddLine(10.0, 10.0, 100.0, 10.0)
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


def profile(grey: np.ndarray, at: int) -> str:
    row = 1 + at * ROW_STEP
    middle = row * ROW_PX
    left, right = COL_PX + 8, COL_PX * (1 + LONG) - 8
    if right > grey.shape[1] or middle + 5 > grey.shape[0]:
        return "off the picture"
    band = grey[max(0, middle - 2) : middle + 6, left:right]
    rows = [255 - float(one.mean()) for one in band]
    return " ".join(f"{one:5.1f}" for one in rows) + f"   ink {sum(rows):6.1f}"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "group.xlsx"
    if not args.reuse:
        build(made)
        if not shoot(made):
            print("  Excel would not hand over a picture")
            return 1
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")).astype(float)
    drawing = dict(os.environ)
    rows = 4 + len(ARMS) * ROW_STEP
    drawing["OXI_XLSX_RANGE"] = f"1,1,{rows},13"
    subprocess.run(
        [str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", env=drawing,
    )
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")).astype(float)
    global COL_PX, ROW_PX
    COL_PX = (truth.shape[1] - 1) // 13
    ROW_PX = (truth.shape[0] - 1) // rows
    print(f"  Excel {truth.shape[1]}x{truth.shape[0]}, Oxi {mine.shape[1]}x{mine.shape[0]}"
          f"; a cell is {COL_PX} x {ROW_PX}")
    print("  8th   ink per row through the rule")
    for at, eighth in enumerate(ARMS):
        print(f"  {eighth:>3} E {profile(truth, at)}")
        print(f"  {'':>3} O {profile(mine, at)}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
