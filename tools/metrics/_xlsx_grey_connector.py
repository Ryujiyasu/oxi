# -*- coding: utf-8 -*-
r"""What makes ONE of `glossary_05`'s connectors soft?

`_xlsx_which_connector.py` named it: the elbow, the fifth connector in the
part. Excel draws it at half ink over two rows where every other connector in
the same file is one solid row, and six things that might explain it have been
measured and falsified on shapes built for the purpose — softness, the theme's
width, a stretched group, the sheet's zoom, the turn itself, and a half-pixel
anchor.

Synthetic arms have run out of ways to be wrong, so this goes the other way:
take the real workbook and take ONE thing away from that shape at a time. The
connector keeps a colour of its own so it can be found in the picture whatever
else changes, and what is reported is how dark its darkest row gets — 0 for a
hard rule, about half for one spread over a boundary.

    python tools\metrics\_xlsx_grey_connector.py
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
SOURCE = (
    REPO / "tools" / "golden-test" / "documents" / "xlsx"
    / "33ac9f9d7afc_20230411_resources_standard_guidelines_glossary_05.xlsx"
)
SCRATCH = Path(r"C:\tmp\xlsx_grey_connector")
PAINT = "00C0C0"
# The elbow, counted among the part's `<xdr:cxnSp>` elements in the order they
# are written. `_xlsx_which_connector.py` reads the same order out.
WHICH = 4


def marked(piece: str) -> str:
    """Give this connector's outline the colour it is found by."""
    def recolour(found: re.Match[str]) -> str:
        return re.sub(
            r"<a:(sysClr|schemeClr|srgbClr)[^>]*(/>|>.*?</a:\1>)",
            f'<a:srgbClr val="{PAINT}"/>',
            found.group(0),
            count=1,
            flags=re.S,
        )

    return re.sub(r"<a:ln\b[^>]*>.*?</a:ln>", recolour, piece, count=1, flags=re.S)


# Each arm: a name, and what it does to the elbow's own XML.
ARMS: list[tuple[str, "object"]] = [
    ("as it stands", lambda one: one),
    ("no stCxn/endCxn", lambda one: re.sub(r"<a:(st|end)Cxn[^>]*/>", "", one)),
    ("no solidFill", lambda one: re.sub(
        r"<a:solidFill>(?:(?!</a:solidFill>).)*?</a:solidFill>(?=<a:ln\b)", "", one, flags=re.S)),
    ("no turn", lambda one: re.sub(r' rot="-?\d+"', "", one)),
    ("no flip", lambda one: re.sub(r' flip[HV]="1"', "", one)),
    ("no turn, no flip", lambda one: re.sub(
        r' (rot="-?\d+"|flip[HV]="1")', "", one)),
    ("no arrowhead", lambda one: re.sub(r"<a:(head|tail)End[^>]*/>", "", one)),
    ("no xdr:style", None),   # handled outside the shape's own element
]


def one_arm(drawing: str, alter, drop_style: bool) -> str:
    out, at = [], 0
    for piece in re.split(r"(<xdr:cxnSp\b.*?</xdr:cxnSp>)", drawing, flags=re.S):
        if not piece.startswith("<xdr:cxnSp"):
            out.append(piece)
            continue
        if at == WHICH:
            piece = marked(alter(piece) if alter else piece)
            if drop_style:
                # The style block sits after the shape, inside the same anchor,
                # so it is taken off the text that follows rather than the
                # shape itself.
                out.append(piece)
                at += 1
                continue
        out.append(piece)
        at += 1
    held = "".join(out)
    if drop_style:
        held = re.sub(
            r"(</xdr:cxnSp>)<xdr:style>.*?</xdr:style>",
            r"\1",
            held,
            flags=re.S,
        )
    return held


def build(made: Path, alter, drop_style: bool) -> None:
    if made.exists():
        made.unlink()
    with zipfile.ZipFile(SOURCE) as was, zipfile.ZipFile(made, "w", zipfile.ZIP_DEFLATED) as now:
        for item in was.infolist():
            held = was.read(item.filename)
            if item.filename == "xl/drawings/drawing1.xml":
                held = one_arm(held.decode("utf-8"), alter, drop_style).encode("utf-8")
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
                time.sleep(0.6)
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


def darkest(picture: Path) -> str:
    """How dark the marked connector gets, and over how many rows.

    Text is the thing to keep out: Excel draws it with subpixel antialiasing,
    so a black glyph leaves pale cyan fringes that answer to any loose test for
    the paint. The connector runs hundreds of pixels across, so only rows
    holding a long run of it are read.
    """
    held = np.asarray(Image.open(picture).convert("RGB")).astype(int)
    lit = (
        (held[:, :, 0] < 200)
        & (abs(held[:, :, 1] - held[:, :, 2]) < 12)
        & (held[:, :, 1] > held[:, :, 0] + 60)
    )
    runs = lit.sum(axis=1)
    rows = [int(y) for y in np.nonzero(runs > 100)[0]]
    if not rows:
        return "no long run of the marked connector in the picture"
    said = []
    for y in rows:
        shade = int(held[y][lit[y]][:, 0].min())
        common = int(np.bincount(held[y][lit[y]][:, 0]).argmax())
        said.append(f"y={y} {common:>3}/{shade:>3} x{int(runs[y]):>4}")
    return "  ".join(said[:4]) + (f"  (+{len(said) - 4} more)" if len(said) > 4 else "")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    SCRATCH.mkdir(parents=True, exist_ok=True)
    for at, (name, alter) in enumerate(ARMS):
        made = SCRATCH / f"arm{at}.xlsx"
        shot = SCRATCH / f"arm{at}.png"
        if not args.reuse:
            build(made, alter, drop_style=(alter is None))
            if not shoot(made, shot):
                print(f"  {name:<20} Excel would not hand over a picture")
                continue
        if not shot.exists():
            print(f"  {name:<20} no picture")
            continue
        print(f"  {name:<20} {darkest(shot)}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
