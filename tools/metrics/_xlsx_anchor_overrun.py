# -*- coding: utf-8 -*-
r"""What does Excel do with an anchor offset that runs past its own cell?

`002` is the corpus floor and its worst tile is a pinned note whose box we draw
seven pixels too tall. Its VML anchor ends at row 3 plus 34 pixels — and row 3
is 27 pixels high, so the offset runs 7 past the cell it is measured into.
Excel draws the note ending exactly at the top of row 4, which is what
CLAMPING the offset to the cell would give.

The same shape of thing showed up while building `_xlsx_bent_connector.py`:
sixteen shapes meant to stand 192pt apart, each written as an offset into
column A, all came back stacked at 48pt — the width of column A.

So the question is whether an offset past its cell is clamped, and it is asked
of Excel directly rather than off a picture: the shapes are written with known
overruns and Excel is asked where it put them and how big they are.

    python tools\metrics\_xlsx_anchor_overrun.py
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
SCRATCH = Path(r"C:\tmp\xlsx_anchor_overrun")

EMU_PX = 9525
ROW_PT = 15.0                       # every row, so the arithmetic is visible
ROW_PX = int(ROW_PT * 96 / 72)      # 20
COL_PT = 48.0
ROW_STEP = 4

# How far past its cell each arm's far corner reaches, in pixels. Zero is the
# control: it must come back as the plain two-cell box or the instrument is
# wrong before the question is asked.
OVERRUNS = [0, 5, 10, 19, 20, 21, 30, 45, 100]
# The same question of the other axis. A column is stated below so its width in
# pixels is known; the far corner reaches this many pixels past it.
ACROSS = [0, 20, 63, 64, 65, 120, 300]


def drawing_xml() -> str:
    shapes = []
    for at, past in enumerate(OVERRUNS):
        row = 1 + at * ROW_STEP
        shapes.append(
            f"<xdr:twoCellAnchor>"
            f"<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            f"<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{row + 1}</xdr:row>"
            f"<xdr:rowOff>{past * EMU_PX}</xdr:rowOff></xdr:to>"
            f"<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr>"
            f"<xdr:cNvPr id=\"{at + 2}\" name=\"box {at}\"/>"
            f"<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>"
            f"<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>"
            f"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
            f"<a:ln w=\"9525\"><a:solidFill>"
            f"<a:srgbClr val=\"000000\"/></a:solidFill></a:ln>"
            f"</xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"
        )
    for at, past in enumerate(ACROSS):
        row = 1 + (len(OVERRUNS) + at) * ROW_STEP
        shapes.append(
            f"<xdr:twoCellAnchor>"
            f"<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            f"<xdr:to><xdr:col>2</xdr:col>"
            f"<xdr:colOff>{past * EMU_PX}</xdr:colOff>"
            f"<xdr:row>{row + 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
            f"<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr>"
            f"<xdr:cNvPr id=\"{100 + at}\" name=\"wide {at}\"/>"
            f"<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>"
            f"<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>"
            f"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
            f"<a:ln w=\"9525\"><a:solidFill>"
            f"<a:srgbClr val=\"000000\"/></a:solidFill></a:ln>"
            f"</xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/'
        'spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/'
        'drawingml/2006/main">' + "".join(shapes) + "</xdr:wsDr>"
    )


def sheet_xml(was: str) -> str:
    """Give every row and column a size of its own, so the sums are readable."""
    rows = "".join(
        f'<row r="{n}" ht="{ROW_PT}" customHeight="1"/>'
        for n in range(1, 4 + (len(OVERRUNS) + len(ACROSS)) * ROW_STEP)
    )
    held = re.sub(r"<sheetData\s*/>|<sheetData>.*?</sheetData>",
                  f"<sheetData>{rows}</sheetData>", was, flags=re.S)
    held = re.sub(r'defaultRowHeight="[\d.]+"', f'defaultRowHeight="{ROW_PT}"', held)
    return held


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
            if item.filename == "xl/worksheets/sheet1.xml":
                held = sheet_xml(held.decode("utf-8")).encode("utf-8")
            now.writestr(item, held)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "overrun.xlsx"
    if not args.reuse:
        build(made)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(made))
    try:
        sheet = book.Worksheets(1)
        rows = sheet.Rows(1).RowHeight
        print(f"  a row is {rows}pt; {sheet.Shapes.Count} shape(s)")
        print(f"  {'past':>5} {'Excel height':>13} {'clamped says':>13}"
              f" {'unclamped says':>15}")
        agree_clamped = agree_not = 0
        for at, past in enumerate(OVERRUNS):
            if at >= sheet.Shapes.Count:
                break
            shape = sheet.Shapes.Item(at + 1)
            tall = round(shape.Height, 2)
            # The box spans one whole row plus whatever the far offset adds.
            clamped = round(ROW_PT + min(past, ROW_PX) * 72 / 96, 2)
            plain = round(ROW_PT + past * 72 / 96, 2)
            agree_clamped += abs(tall - clamped) < 0.51
            agree_not += abs(tall - plain) < 0.51
            print(f"  {past:>5} {tall:>13} {clamped:>13} {plain:>15}"
                  f"   {'clamped' if abs(tall - clamped) < 0.51 else ''}")
        print(f"  clamped fits {agree_clamped} of {len(OVERRUNS)};"
              f" unclamped fits {agree_not}")
        wide = round(sheet.Columns(3).Width, 2)
        print(f"  column C is {wide}pt")
        print(f"  {'past':>5} {'Excel width':>13} {'clamped says':>13}"
              f" {'unclamped says':>15}")
        fits = 0
        for at, past in enumerate(ACROSS):
            seat = len(OVERRUNS) + at
            if seat >= sheet.Shapes.Count:
                break
            shape = sheet.Shapes.Item(seat + 1)
            across = round(shape.Width, 2)
            room = wide * 96 / 72
            clamped = round(wide + min(past, room) * 72 / 96, 2)
            plain = round(wide + past * 72 / 96, 2)
            fits += abs(across - clamped) < 0.76
            print(f"  {past:>5} {across:>13} {clamped:>13} {plain:>15}"
                  f"   {'clamped' if abs(across - clamped) < 0.76 else ''}")
        print(f"  across, clamped fits {fits} of {len(ACROSS)}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
