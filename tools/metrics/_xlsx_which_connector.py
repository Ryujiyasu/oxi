# -*- coding: utf-8 -*-
r"""Which connector is the grey one?

`glossary_05` draws some of its flowchart connectors as two rows of exactly
127 — one pixel of ink spread across a boundary — and others as solid black.
Four hypotheses for the difference have been measured and falsified, and the
next step needs the actual shape rather than another synthetic one: which
`<xdr:twoCellAnchor>` of the file is the grey line at y=659?

Deleting shapes one at a time would take one Excel picture each. Instead every
connector is given a colour of its own in a copy of the workbook, so ONE
picture names them all: read the hue at a row and the anchor is known.

    python tools\metrics\_xlsx_which_connector.py
    python tools\metrics\_xlsx_which_connector.py --at 659
"""

from __future__ import annotations

import argparse
import re
import shutil
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
SCRATCH = Path(r"C:\tmp\xlsx_which_connector")

# Colours far enough apart to tell one from another after Excel has softened
# an edge: each is a primary or a pair of them at full strength, so a half-lit
# pixel of any of them still says which it came from.
PAINTS = [
    "FF0000", "00C000", "0000FF", "FF00FF", "00C0C0", "C08000", "800080",
    "008000", "804000", "0080FF", "FF8080", "80FF00", "FF0080", "00FF80",
    "8080FF", "C0C000", "406080", "804080", "408040", "606060",
]


def painted(drawing: str) -> tuple[str, list[str]]:
    """Give every connector's outline a colour of its own."""
    out = []
    order = []
    at = 0
    for piece in re.split(r"(<xdr:cxnSp\b.*?</xdr:cxnSp>)", drawing, flags=re.S):
        if not piece.startswith("<xdr:cxnSp"):
            out.append(piece)
            continue
        geometry = re.search(r'prst="([A-Za-z0-9]+)"', piece)
        paint = PAINTS[at % len(PAINTS)]
        order.append(f"{at}:{geometry.group(1) if geometry else '?'}:{paint}")
        # Only the outline's own fill, which is the one inside `<a:ln>`.
        def recolour(found: re.Match[str]) -> str:
            body = found.group(0)
            body = re.sub(
                r"<a:(sysClr|schemeClr|srgbClr)[^>]*(/>|>.*?</a:\1>)",
                f'<a:srgbClr val="{paint}"/>',
                body,
                count=1,
                flags=re.S,
            )
            return body

        piece = re.sub(r"<a:ln\b[^>]*>.*?</a:ln>", recolour, piece, count=1, flags=re.S)
        out.append(piece)
        at += 1
    return "".join(out), order


def build(made: Path) -> list[str]:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    order: list[str] = []
    if made.exists():
        made.unlink()
    with zipfile.ZipFile(SOURCE) as was, zipfile.ZipFile(made, "w", zipfile.ZIP_DEFLATED) as now:
        for item in was.infolist():
            held = was.read(item.filename)
            if item.filename == "xl/drawings/drawing1.xml":
                said, order = painted(held.decode("utf-8"))
                held = said.encode("utf-8")
            now.writestr(item, held)
    return order


def shoot(made: Path) -> bool:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(made))
    try:
        sheet = book.Worksheets(1)
        used = sheet.UsedRange
        for _ in range(10):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(1.2)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    parser.add_argument("--at", type=int, action="append",
                        help="a row to name the colours on (repeatable)")
    args = parser.parse_args()
    made = SCRATCH / "painted.xlsx"
    order = build(made) if not args.reuse else []
    if not args.reuse and not shoot(made):
        print("  Excel would not hand over a picture")
        return 1
    for one in order:
        print(f"  {one}")
    picture = np.asarray(Image.open(SCRATCH / "excel.png").convert("RGB")).astype(int)
    print(f"  picture {picture.shape[1]}x{picture.shape[0]}")
    for row in args.at or [659]:
        if row >= picture.shape[0]:
            print(f"  row {row} is off the picture")
            continue
        seen: dict[tuple[int, int, int], int] = {}
        for x in range(picture.shape[1]):
            for y in (row, row + 1):
                if y >= picture.shape[0]:
                    continue
                pixel = tuple(int(one) for one in picture[y, x])
                if pixel == (255, 255, 255):
                    continue
                seen[pixel] = seen.get(pixel, 0) + 1
        held = sorted(seen.items(), key=lambda one: -one[1])[:6]
        print(f"  row {row}: " + "  ".join(f"{p}x{n}" for p, n in held))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
