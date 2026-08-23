"""What do a `rightBrace`'s two adjusts do to its outline?

`_xlsx_brace_shape.py` read the shape at its defaults: a quarter-ellipse
corner whose radius is `min(w,h) x 0.08333` down and `w/2` across, a straight
body at `x = w/2`, and the point at `(w, h/2)`. But the corpus does not use
the defaults — its 24 braces carry adj1 between 17244 and 58333 and adj2
between 11152 and 50000 — so both have to be swept before any of it is built.

The adjusts are written into the drawing by hand rather than set through COM,
whose `Adjustments` are in units of their own; the file is then reopened and
Excel asked to draw it. Read back per arm: how far down the corner reaches
the body, and which row the point sits on.

Run: python tools/metrics/_xlsx_brace_adjust.py
"""

from __future__ import annotations

import re
import sys
import time
import zipfile
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_brace_adjust")
RIGHT_BRACE = 32
TOP_PT = 40.0
LEFT_PT = 60.0
WIDE_PT, HIGH_PT = 40.0, 300.0
# The defaults, the three pairs the corpus actually carries, and two more.
ARMS = [(8333, 50000), (58333, 11152), (17244, 37031), (17244, 50000),
        (30000, 25000), (5000, 75000)]


def seed() -> Path:
    """One workbook holding one brace, for the adjusts to be written into."""
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "seed.xlsx"
    if made.exists():
        made.unlink()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:Z60").Interior.Color = 0xFFFFFF
        shape = sheet.Shapes.AddShape(RIGHT_BRACE, LEFT_PT, TOP_PT, WIDE_PT, HIGH_PT)
        shape.Fill.Visible = False
        shape.Line.ForeColor.RGB = 0x000000
        shape.Line.Weight = 0.75
        book.SaveAs(str(made), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return made


def with_adjusts(source: Path, adj1: int, adj2: int) -> Path:
    out = SCRATCH / f"brace_{adj1}_{adj2}.xlsx"
    if out.exists():
        out.unlink()
    held = zipfile.ZipFile(source)
    avlst = (f'<a:avLst><a:gd name="adj1" fmla="val {adj1}"/>'
             f'<a:gd name="adj2" fmla="val {adj2}"/></a:avLst>')
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as writing:
        for item in held.infolist():
            raw = held.read(item.filename)
            if item.filename.startswith("xl/drawings/") and item.filename.endswith(".xml"):
                text = raw.decode("utf-8")
                text = re.sub(
                    r'(<a:prstGeom prst="rightBrace">)<a:avLst/>',
                    lambda m: m.group(1) + avlst,
                    text,
                )
                raw = text.encode("utf-8")
            writing.writestr(item, raw)
    return out


def picture(sheet):
    for _ in range(8):
        try:
            sheet.Activate()
            sheet.Range("A1:Z60").CopyPicture(Appearance=1, Format=2)
        except Exception:
            time.sleep(0.6)
            continue
        time.sleep(0.5)
        held = ImageGrab.grabclipboard()
        if held is not None:
            return held
    return None


def fit_corner(image, top: int, left: int, wide: int, rows: int = 30) -> float | None:
    """The corner's y-radius, fitted from the outline rather than eyeballed.

    The corner is a quarter ellipse `x = (w/2) sin t, y = y1 (1 - cos t)`, so
    every row of it gives `y1 = y / (1 - sqrt(1 - (2x/w)^2))`. Averaging those
    beats asking "which row does it look vertical on", which reads low — and
    reads lower the larger the radius, because a shallow approach reaches
    within a pixel of the body sooner.
    """
    grey = np.asarray(image.convert("L"))[top - 2:top + rows, left - 2:left + wide + 3]
    ink = grey < 140
    found = []
    for y in range(2, ink.shape[0]):
        lit = np.where(ink[y])[0]
        if not len(lit):
            continue
        middle = ((int(lit.min()) + int(lit.max())) / 2) - 2
        share = 2 * middle / wide
        if not 0.25 < share < 0.9:
            continue
        under = 1 - share * share
        if under <= 0:
            continue
        found.append((y - 2) / (1 - under ** 0.5))
    if not found:
        return None
    found.sort()
    return found[len(found) // 2]


def read(image, top: int, left: int, high: int, wide: int):
    """Where the corner meets the body, and which row the point is on."""
    grey = np.asarray(image.convert("L"))[top - 2:top + high + 3, left - 2:left + wide + 3]
    ink = grey < 140
    body, point, reach = None, None, -1
    for y in range(ink.shape[0]):
        lit = np.where(ink[y])[0]
        if not len(lit):
            continue
        low, far = int(lit.min()) - 2, int(lit.max()) - 2
        if body is None and y > 3 and low >= wide // 2 - 1 and far <= wide // 2 + 1:
            body = y - 2
        if far > reach:
            reach, point = far, y - 2
    return body, point, reach


def main() -> int:
    source = seed()
    wide, high = round(WIDE_PT * 96 / 72), round(HIGH_PT * 96 / 72)
    smaller = min(wide, high)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    print(f"box {wide}x{high}px, min(w,h) = {smaller}")
    print("  adj1    adj2     corner  point      says corner  says point")
    try:
        for adj1, adj2 in ARMS:
            book_path = with_adjusts(source, adj1, adj2)
            book = excel.Workbooks.Open(str(book_path))
            try:
                held = picture(book.Worksheets(1))
                if held is None:
                    print(f"  {adj1:<7} {adj2:<7}  Excel would not hand over a picture")
                    continue
                held.save(SCRATCH / f"brace_{adj1}_{adj2}.png")
                body, point, _ = read(
                    held, round(TOP_PT * 96 / 72), round(LEFT_PT * 96 / 72), high, wide
                )
                # The preset as remembered: adj2 pinned to 0..100000; adj1
                # capped by half the smaller of adj2 and its complement, scaled
                # by h/ss; corner = ss x a1, point = h x a2.
                a2 = min(max(adj2, 0), 100000)
                cap = min(100000 - a2, a2) / 2 * high / smaller
                a1 = min(max(adj1, 0), cap)
                says_corner = smaller * a1 / 100000
                says_point = high * a2 / 100000
                fitted = fit_corner(held, round(TOP_PT * 96 / 72),
                                    round(LEFT_PT * 96 / 72), wide)
                shown = f"{fitted:5.1f}" if fitted else "  n/a"
                print(f"  {adj1:<7} {adj2:<7}  {str(body):>5}  {str(point):>5}"
                      f"   fitted y1 {shown}   says {says_corner:6.1f} / point {says_point:7.1f}")
            finally:
                book.Close(SaveChanges=False)
    finally:
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
