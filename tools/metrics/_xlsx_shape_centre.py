# -*- coding: utf-8 -*-
r"""How does Excel divide what is left over when a block is centred?

The renderer puts a centred block at `area top + round(slack / 2)`, with
`slack` the room the box has left after the block's own (fractional) height.
`glossary_05`'s flowchart is eight centred boxes out of twelve and four of
them still sit a pixel or two from Excel's, so the division is worth asking
about rather than assuming.

The area is a whole-pixel rectangle, so the slack moves a pixel at a time and
half of it alternates between a whole number and a half. Sweeping the box's
height one pixel at a time therefore walks the rounding boundary again and
again: over sixteen heights the three candidate rules give three different
sequences.

Three faces at one size, whose own line heights end in .666, .200 and .800, so
the sweep meets the boundary in a different phase in each.

    python tools\metrics\_xlsx_shape_centre.py
    python tools\metrics\_xlsx_shape_centre.py --reuse
"""
import argparse
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import numpy as np
from PIL import Image

REPO = Path(__file__).resolve().parents[2]
SHOOTER = Path(__file__).resolve().parent / "_xlsx_screen_shot.ps1"
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SCRATCH = Path(r"C:\tmp\xlsx_shape_centre")
BOOK = SCRATCH / "centre.xlsx"
EMU = 9525.0

SPACING = 5             # rows a case
WORD = "A"              # sits on the baseline, so its ink foot is baseline - 1
FACES = [("Yu Gothic UI", 12.0), ("メイリオ", 12.0), ("ＭＳ Ｐゴシック", 12.0)]
HEIGHTS = list(range(40, 56))


def cases():
    return [(face, size, tall) for face, size in FACES for tall in HEIGHTS]


def anchors_xml():
    held = []
    for index, (face, points, tall) in enumerate(cases()):
        one = (
            f'<a:p><a:pPr algn="l"/>'
            f'<a:r><a:rPr lang="ja-JP" sz="{int(points * 100)}">'
            f'<a:latin typeface="{face}"/><a:ea typeface="{face}"/>'
            f"</a:rPr><a:t>{WORD}</a:t></a:r></a:p>"
        )
        held.append(
            f"<xdr:oneCellAnchor>"
            f"<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{index * SPACING}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            f'<xdr:ext cx="{int(200 * EMU)}" cy="{int(tall * EMU)}"/>'
            f'<xdr:sp macro="" textlink="">'
            f'<xdr:nvSpPr><xdr:cNvPr id="{index + 2}" name="box {index}"/>'
            f"<xdr:cNvSpPr/></xdr:nvSpPr>"
            f'<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f"<a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>"
            # Centred, which is the whole question.
            f'<xdr:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>'
            f"{one}</xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>"
        )
    return "".join(held)


def build():
    from openpyxl import Workbook

    SCRATCH.mkdir(parents=True, exist_ok=True)
    plain = SCRATCH / "_plain.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["A"].width = 2.0
    sheet.column_dimensions["B"].width = 40.0
    sheet.cell(row=1, column=1, value="a")
    sheet.cell(row=len(cases()) * SPACING + 2, column=4, value="z")
    book.save(plain)

    drawing = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f"{anchors_xml()}</xdr:wsDr>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"'
        ' Target="../drawings/drawing1.xml"/></Relationships>'
    )
    BOOK.unlink(missing_ok=True)
    with zipfile.ZipFile(plain) as source, \
            zipfile.ZipFile(BOOK, "w", zipfile.ZIP_DEFLATED) as out:
        for item in source.infolist():
            held = source.read(item.filename)
            if item.filename == "[Content_Types].xml":
                held = held.decode("utf-8").replace(
                    "</Types>",
                    '<Override PartName="/xl/drawings/drawing1.xml"'
                    ' ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                    "</Types>",
                ).encode("utf-8")
            if item.filename == "xl/worksheets/sheet1.xml":
                held = held.decode("utf-8").replace(
                    "</worksheet>", '<drawing r:id="rId1"/></worksheet>')
                if "xmlns:r=" not in held:
                    held = held.replace(
                        "<worksheet ",
                        '<worksheet xmlns:r="http://schemas.openxmlformats.org/'
                        'officeDocument/2006/relationships" ', 1)
                held = held.encode("utf-8")
            out.writestr(item, held)
        out.writestr("xl/drawings/drawing1.xml", drawing)
        out.writestr("xl/worksheets/_rels/sheet1.xml.rels", rels)


def shoot():
    picture = BOOK.with_suffix(".excel.png")
    picture.unlink(missing_ok=True)
    listing = SCRATCH / "_batch.txt"
    listing.write_text(f"{BOOK.resolve()}\t{picture.resolve()}", encoding="utf-8-sig")
    subprocess.run(["powershell", "-NoProfile", "-File", str(SHOOTER),
                    "-ListFile", str(listing.resolve())],
                   capture_output=True, text=True, encoding="utf-8",
                   errors="replace", timeout=1800)
    listing.unlink(missing_ok=True)
    return picture


def drawn():
    ours = SCRATCH / "centre.oxi.png"
    done = subprocess.run(
        [str(RENDERER), str(BOOK), str(ours), "96"], capture_output=True, timeout=1800,
        env=dict(os.environ, OXI_XLSX_DUMP_ROWS="1", OXI_XLSX_DUMP_COLUMNS="1",
                 OXI_XLSX_DUMP_BLOCK="1", OXI_XLSX_SHAPE_TEXT="1"))
    heights, lane = {}, 0
    for line in done.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            heights[int(parts[1])] = int(float(parts[3]))
        if len(parts) == 4 and parts[0] == "column" and lane == 0:
            lane = int(float(parts[3]))
    told = []
    for line in done.stderr.decode("utf-8", "replace").splitlines():
        found = re.search(r"^block area (\d+)\.\.(\d+).*?block=([0-9.]+).*? at=([0-9.]+)", line)
        if found:
            told.append(tuple(float(one) for one in found.groups()))
    return ours, heights, lane, told


def foot(picture, top, bottom, lane):
    """The last row of ink — for `A`, one above the baseline."""
    held = (picture[top:bottom, lane:] < 128).sum(axis=1)
    lit = np.where(held > 0)[0]
    return top + int(lit[-1]) if len(lit) else None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")

    build()
    picture = BOOK.with_suffix(".excel.png") if args.reuse else shoot()
    if not picture.exists():
        print("Excel gave no picture")
        return
    truth = np.asarray(Image.open(picture).convert("L"))
    ours_png, heights, lane, told = drawn()
    mine = np.asarray(Image.open(ours_png).convert("L"))
    edges, at = {}, 0
    for index in sorted(heights):
        edges[index] = at
        at += heights[index]

    print(f"{'face':<14}{'box':>5}{'area':>12}{'block':>9}{'slack':>8}"
          f"{'Excel':>7}{'ours':>6}{'round':>7}{'floor':>7}{'ceil':>6}")
    seen = {}
    for index, (face, points, tall) in enumerate(cases()):
        top = edges.get(index * SPACING)
        bottom = edges.get(index * SPACING + SPACING - 1)
        if top is None or bottom is None or index >= len(told):
            continue
        if bottom > min(truth.shape[0], mine.shape[0]):
            continue
        theirs = foot(truth, top, bottom, lane)
        mine_foot = foot(mine, top, bottom, lane)
        if theirs is None or mine_foot is None:
            continue
        area_top, area_foot, block, _at = told[index]
        slack = (area_foot - area_top) - block
        # What each rule would put the block's top at, and so — since the
        # leading is the same in all three — how the ink foot would move.
        base = {name: area_top + rule(slack / 2.0)
                for name, rule in (("round", lambda v: np.floor(v + 0.5)),
                                   ("floor", np.floor), ("ceil", np.ceil))}
        shown = {name: theirs - (mine_foot - (_at - value))
                 for name, value in base.items()}
        got = [name for name, value in base.items() if value == _at]
        seen.setdefault(face, []).append(
            (tall, theirs - mine_foot, {name: int(value - _at) for name, value in base.items()}))
        print(f"{face:<14}{tall:>5}{f'{area_top:.0f}..{area_foot:.0f}':>12}{block:>9.3f}"
              f"{slack:>8.3f}{theirs:>7}{mine_foot:>6}"
              f"{int(base['round'] - _at):>7}{int(base['floor'] - _at):>7}"
              f"{int(base['ceil'] - _at):>6}"
              f"{'' if theirs == mine_foot else '  <<'}")
    print("\nExcel less ours, by face — a rule fits when its column below "
          "matches this everywhere")
    for face, rows in seen.items():
        print(f"  {face:<14}{[one[1] for one in rows]}")
        for name in ("round", "floor", "ceil"):
            print(f"    {name:<12}{[one[2][name] for one in rows]}")


if __name__ == "__main__":
    main()
