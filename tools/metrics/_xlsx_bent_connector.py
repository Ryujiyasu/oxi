# -*- coding: utf-8 -*-
r"""Where does an elbow connector's corner go once the shape is turned?

`glossary_05` is the corpus floor, and one of the two things missing from its
flowchart is a single `bentConnector2` — an L-shaped connector, turned a
quarter turn and mirrored. The renderer draws nothing for a geometry it does
not know, so the L is simply absent, and nothing in the xlsx path reads `rot`
at all.

Neither half can be taken from the format's own words alone. The preset says
the path runs from one corner across and then down; which corner that ends up
being depends on the flips and on the turn, and the turn is about the box's
centre, so a tall box drawn a quarter turn round is a wide one. So this asks
Excel: sixteen arms, every turn against every flip, one connector an arm.

What comes back for each arm is where the ink actually is — the row and column
the two arms of the L lie on, and how far each one reaches — read the same way
out of Excel's picture and out of ours.

    python tools\metrics\_xlsx_bent_connector.py
    python tools\metrics\_xlsx_bent_connector.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_bent_connector")

EMU = 12700          # per point
EMU_PX = 9525        # per pixel at 96dpi
TURNS = [0, 5400000, 10800000, 16200000]     # 0, 90, 180, 270 degrees
FLIPS = [(False, False), (True, False), (False, True), (True, True)]
ARMS = [(turn, flip) for turn in TURNS for flip in FLIPS]

# One arm's box, and the pitch it is laid out on. The box is deliberately not
# square: a quarter turn swaps its sides, and a square one would hide that.
WIDE, TALL = 96.0, 60.0
COLUMNS = 4
# Each arm hangs from a cell of its own with no offset. An offset was tried
# first and is a trap: Excel CLAMPS `colOff` to the width of the column it is
# an offset into, so four arms meant to stand 192pt apart all came back at
# 48pt, stacked on top of one another. Cells are the only way to place a
# drawing far from the origin.
COL_STEP, ROW_STEP = 4, 8
# What one cell of the untouched sheet comes to. Both pictures are 1985 x 1201
# for the 31 columns and 60 rows copied, which is where these come from; the
# run says so again, so a change in either is seen rather than assumed.
COL_PX, ROW_PX = 64, 20
# How far into its row each arm starts, in EMU. Zero puts every corner on a
# whole pixel, which is the one place a soft edge and a hard one look the
# same; half a pixel is what tells them apart.
DOWN = 0


def cell_of(at: int) -> tuple[int, int]:
    # One cell in from the corner. An arm standing ON the first row or column
    # has an edge of its L at the very first pixel of the picture, where the
    # sheet's own border sits over it: seven arms read as half-drawn until they
    # were moved off it, and the half missing was always the one at the edge.
    return (1 + (at % COLUMNS) * COL_STEP, 1 + (at // COLUMNS) * ROW_STEP)


def placed(at: int) -> tuple[float, float]:
    """Where an arm's box starts, in points."""
    col, row = cell_of(at)
    return (col * COL_PX * 72 / 96, row * ROW_PX * 72 / 96)


def drawing_xml() -> str:
    shapes = []
    for at, (turn, (flip_h, flip_v)) in enumerate(ARMS):
        col, row = cell_of(at)
        turned = f' rot="{turn}"' if turn else ""
        mirror = (' flipH="1"' if flip_h else "") + (' flipV="1"' if flip_v else "")
        # Hung from the first cell with the whole position as its offset, which
        # is absolute placement in the one form the corpus actually uses: not
        # one of its 2329 anchors is an `absoluteAnchor`.
        shapes.append(
            f"<xdr:oneCellAnchor>"
            f"<xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{row}</xdr:row><xdr:rowOff>{DOWN}</xdr:rowOff></xdr:from>"
            f"<xdr:ext cx=\"{int(WIDE * EMU)}\" cy=\"{int(TALL * EMU)}\"/>"
            f"<xdr:cxnSp macro=\"\"><xdr:nvCxnSpPr>"
            f"<xdr:cNvPr id=\"{at + 2}\" name=\"bent {at}\"/>"
            f"<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>"
            f"<a:xfrm{turned}{mirror}>"
            f"<a:off x=\"0\" y=\"0\"/>"
            f"<a:ext cx=\"{int(WIDE * EMU)}\" cy=\"{int(TALL * EMU)}\"/></a:xfrm>"
            f"<a:prstGeom prst=\"bentConnector2\"><a:avLst/></a:prstGeom>"
            f"<a:ln w=\"9525\"><a:solidFill>"
            f"<a:srgbClr val=\"000000\"/></a:solidFill></a:ln>"
            f"</xdr:spPr></xdr:cxnSp><xdr:clientData/></xdr:oneCellAnchor>"
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/'
        'spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/'
        'drawingml/2006/main">' + "".join(shapes) + "</xdr:wsDr>"
    )


def build(made: Path) -> None:
    """A workbook holding nothing but the connectors, written by hand.

    Excel's own object model has no way to ask for this shape at a stated turn
    — an elbow connector it draws for itself is a `bentConnector3` — so the
    arms are written into the drawing part directly and Excel is asked to
    render what it reads.
    """
    SCRATCH.mkdir(parents=True, exist_ok=True)
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<dimension ref="A1:AE60"/><sheetFormatPr defaultRowHeight="15"/><sheetData/><drawing r:id="rId1"/></worksheet>'
    )
    with zipfile.ZipFile(made, "w", zipfile.ZIP_DEFLATED) as zip_:
        zip_.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            "</Types>",
        )
        zip_.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>",
        )
        zip_.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="bent" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        zip_.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            "</Relationships>",
        )
        zip_.writestr(
            "xl/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
            '<cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>',
        )
        zip_.writestr("xl/worksheets/sheet1.xml", sheet)
        zip_.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
            "</Relationships>",
        )
        zip_.writestr("xl/drawings/drawing1.xml", drawing_xml())


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
                sheet.Range("A1:AE60").CopyPicture(Appearance=1, Format=2)
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


def limb(dark: np.ndarray, at: int, grey: np.ndarray | None = None) -> str:
    """The two arms of one L, read out of the picture as rows and columns.

    A quarter-turn box is a different shape from the one the anchor states, so
    the window read is generous — the pitch either way, from the box's own
    corner — and what is reported is where the ink lies inside it rather than
    whether it fills it.
    """
    x, y = placed(at)
    left, top = int(x * 96 / 72) - 24, int(y * 96 / 72) - 24
    right = min(dark.shape[1], left + COL_STEP * COL_PX)
    bottom = min(dark.shape[0], top + ROW_STEP * ROW_PX)
    left, top = max(0, left), max(0, top)
    window = dark[top:bottom, left:right]
    if not window.any():
        return "nothing drawn"
    rows = window.sum(axis=1)
    cols = window.sum(axis=0)
    across = int(rows.argmax())
    down = int(cols.argmax())
    lit_x = np.nonzero(window[across])[0]
    lit_y = np.nonzero(window[:, down])[0]
    # How dark the darkest row of the arm is. A hard rule reaches 0; a rule
    # spread over a boundary stops at about half, which is the whole question.
    shade = int(grey[top:bottom, left:right][across].min()) if grey is not None else -1
    return (
        f"y={across + top - int(y * 96 / 72):>4} x {lit_x[0] + left - int(x * 96 / 72):>4}"
        f"..{lit_x[-1] + left - int(x * 96 / 72):>4}   "
        f"x={down + left - int(x * 96 / 72):>4} y {lit_y[0] + top - int(y * 96 / 72):>4}"
        f"..{lit_y[-1] + top - int(y * 96 / 72):>4}  ink {shade:>3}"
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    parser.add_argument("--half", action="store_true",
                        help="start every arm half a pixel into its row")
    args = parser.parse_args()
    global DOWN
    if args.half:
        DOWN = EMU_PX // 2
    made = SCRATCH / "bent.xlsx"
    if not args.reuse:
        build(made)
        if not shoot(made):
            print("  Excel would not hand over a picture")
            return 1
    truth_grey = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")).astype(int)
    truth = truth_grey < 140
    drawing = dict(os.environ)
    drawing["OXI_XLSX_RANGE"] = "1,1,60,31"
    subprocess.run(
        [str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", env=drawing,
    )
    mine_grey = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")).astype(int)
    mine = mine_grey < 140
    print(f"  Excel {truth.shape[1]}x{truth.shape[0]}, Oxi {mine.shape[1]}x{mine.shape[0]}")
    print(f"  {'turn':>5} {'flip':<5} {'Excel':<44}{'Oxi'}")
    agree = 0
    for at, (turn, (flip_h, flip_v)) in enumerate(ARMS):
        flip = ("H" if flip_h else "") + ("V" if flip_v else "") or "-"
        one, two = limb(truth, at, truth_grey), limb(mine, at, mine_grey)
        same = one == two
        agree += same
        print(f"  {turn // 60000:>5} {flip:<5} {one:<44}{two}{'' if same else '  <<'}")
    print(f"  {agree} of {len(ARMS)} arms agree")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
