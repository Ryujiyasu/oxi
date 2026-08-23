"""What shape is a `bracketPair`, in pixels?

Five of them in the corpus, four on a drawn sheet, and the renderer has no arm
for the preset. They are wide and shallow — 223x35 and 374x33 — with `adj`
7853 and 16667, so the corner is only a pixel or three and has to be read
carefully rather than guessed.

Same method as the brace: the adjust is written into the drawing by hand, the
file reopened, and the outline read off Excel's own picture — for every row,
the leftmost run of ink and the rightmost, since a bracket pair is two strokes
and a per-row min/max would span the gap between them.

Run: python tools/metrics/_xlsx_bracket_shape.py
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

SCRATCH = Path(r"C:\tmp\xlsx_bracket")
# Found by adding every id 1..139 and reading the `prst` each one wrote:
# bevel 15, bracketPair 26, leftBrace 31, rightBrace 32, upArrow 35,
# uturnArrow 42. Guessing the number is how this first drew a `rightArrow`
# and read an outline off it.
BRACKET_PAIR = 26
TOP_PT = 40.0
LEFT_PT = 60.0
ARMS = [
    ((120.0, 90.0), 16667),
    ((120.0, 90.0), 7853),
    ((120.0, 90.0), 40000),
    ((240.0, 30.0), 7853),
    ((60.0, 200.0), 25000),
]


def seed(wide_pt: float, high_pt: float) -> Path:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / f"seed_{wide_pt:.0f}x{high_pt:.0f}.xlsx"
    if made.exists():
        return made
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:Z60").Interior.Color = 0xFFFFFF
        shape = sheet.Shapes.AddShape(BRACKET_PAIR, LEFT_PT, TOP_PT, wide_pt, high_pt)
        shape.Fill.Visible = False
        shape.Line.ForeColor.RGB = 0x000000
        shape.Line.Weight = 0.75
        book.SaveAs(str(made), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return made


def with_adjust(source: Path, adj: int) -> tuple[Path, str]:
    out = SCRATCH / f"{source.stem}_{adj}.xlsx"
    if out.exists():
        out.unlink()
    held = zipfile.ZipFile(source)
    named = ""
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as writing:
        for item in held.infolist():
            raw = held.read(item.filename)
            if item.filename.startswith("xl/drawings/") and item.filename.endswith(".xml"):
                text = raw.decode("utf-8")
                found = re.search(r'<a:prstGeom prst="([^"]+)"', text)
                named = found.group(1) if found else ""
                text = re.sub(
                    r'(<a:prstGeom prst="[^"]+">)<a:avLst/>',
                    lambda m: m.group(1)
                    + f'<a:avLst><a:gd name="adj" fmla="val {adj}"/></a:avLst>',
                    text,
                )
                raw = text.encode("utf-8")
            writing.writestr(item, raw)
    return out, named


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


def strokes(image, top: int, left: int, high: int, wide: int):
    """Per row, the centre of the leftmost run of ink and of the rightmost."""
    grey = np.asarray(image.convert("L"))[top - 3:top + high + 4, left - 3:left + wide + 4]
    ink = grey < 140
    out = {}
    for y in range(ink.shape[0]):
        lit = np.where(ink[y])[0]
        if not len(lit):
            continue
        runs, start, last = [], int(lit[0]), int(lit[0])
        for x in lit[1:]:
            if int(x) - last > 1:
                runs.append((start, last))
                start = int(x)
            last = int(x)
        runs.append((start, last))
        out[y - 3] = (
            (runs[0][0] + runs[0][1]) / 2 - 3,
            (runs[-1][0] + runs[-1][1]) / 2 - 3,
            len(runs),
        )
    return out


def main() -> int:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        for (wide_pt, high_pt), adj in ARMS:
            source = seed(wide_pt, high_pt)
            book_path, named = with_adjust(source, adj)
            wide, high = round(wide_pt * 96 / 72), round(high_pt * 96 / 72)
            smaller = min(wide, high)
            book = excel.Workbooks.Open(str(book_path))
            try:
                held = picture(book.Worksheets(1))
                if held is None:
                    print(f"  {wide}x{high} adj {adj}: no picture")
                    continue
                held.save(SCRATCH / f"{book_path.stem}.png")
                rows = strokes(held, round(TOP_PT * 96 / 72), round(LEFT_PT * 96 / 72),
                               high, wide)
                lit = sorted(rows)
                says = smaller * adj / 100_000
                print(f"  preset {named}  box {wide}x{high}  adj {adj}"
                      f"   corner should be {says:.1f}px")
                # Where the arc ends and the straight side begins, which is
                # the corner radius itself.
                flat = [y for y in lit if rows[y][0] <= 0.6]
                print(f"    the side is straight from y={flat[0]} to y={flat[-1]}"
                      f"   (so the corner reads {flat[0]:.0f} and {high - flat[-1]:.0f})")
                # And the arc's own shape, against a circle of that radius.
                import math
                for y in lit[:6]:
                    seen = rows[y][0]
                    inside = says * says - (says - y) ** 2
                    circle = says - math.sqrt(inside) if inside > 0 else says
                    print(f"    y {y:3d}: left {seen:6.1f}   a circle of r={says:.1f} says {circle:6.1f}")
            finally:
                book.Close(SaveChanges=False)
    finally:
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
