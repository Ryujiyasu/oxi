"""What shape is a `rightBrace`, in pixels?

`b6a3a84180c9_002` — the lowest-scoring workbook — draws a large curly brace
beside its notes, and Oxi draws nothing there: `rightBrace` is a preset the
renderer has no arm for. The corpus holds 24 of them across three workbooks.

Rather than trust a remembered preset definition, this reads the shape off
Excel's own picture: one brace per size, no fill, a hairline rule, and for
every row of the box the leftmost and rightmost lit pixel. What comes back is
the outline itself, which the implementation can then be held against.

Run: python tools/metrics/_xlsx_brace_shape.py
"""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_brace")
# msoShapeRightBrace. 92 is a different shape entirely — it drew a thin spike
# and the outline read off it was nonsense, so the workbook is saved once and
# the preset name checked against the XML before any of it is believed.
RIGHT_BRACE = 32
LEFT_BRACE = 31
TOP_PT = 40.0
LEFT_PT = 60.0
SIZES = [(30.0, 300.0), (16.0, 520.0), (60.0, 160.0), (40.0, 40.0)]


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


def outline(image, top: int, left: int, high: int, wide: int):
    """For each row of the box, the leftmost and rightmost lit pixel."""
    grey = np.asarray(image.convert("L"))[top - 2:top + high + 2, left - 2:left + wide + 2]
    ink = grey < 140
    rows = []
    for y in range(ink.shape[0]):
        lit = np.where(ink[y])[0]
        rows.append((y - 2, int(lit.min()) - 2, int(lit.max()) - 2) if len(lit) else (y - 2, None, None))
    return rows


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    held_all = {}
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:Z60").Interior.Color = 0xFFFFFF
        for wide_pt, high_pt in SIZES:
            shape = sheet.Shapes.AddShape(RIGHT_BRACE, LEFT_PT, TOP_PT, wide_pt, high_pt)
            try:
                shape.Fill.Visible = False
                shape.Line.ForeColor.RGB = 0x000000
                shape.Line.Weight = 0.75
                held = picture(sheet)
                if held is None:
                    print(f"  {wide_pt}x{high_pt}: Excel would not hand over a picture")
                    continue
                held.save(SCRATCH / f"brace_{wide_pt:.0f}x{high_pt:.0f}.png")
                top = round(TOP_PT * 96 / 72)
                left = round(LEFT_PT * 96 / 72)
                wide = round(wide_pt * 96 / 72)
                high = round(high_pt * 96 / 72)
                rows = outline(held, top, left, high, wide)
                held_all[f"{wide_pt:.0f}x{high_pt:.0f}"] = rows
                lit = [r for r in rows if r[1] is not None]
                print(f"  box {wide}x{high}px — ink on {len(lit)} of {len(rows)} rows")
                if lit:
                    print(f"    top row    {lit[0]}")
                    quarter = lit[len(lit) // 4]
                    middle = lit[len(lit) // 2]
                    print(f"    quarter    {quarter}")
                    print(f"    middle     {middle}")
                    print(f"    bottom row {lit[-1]}")
                    reach = max(r[2] for r in lit)
                    where = [r[0] for r in lit if r[2] == reach]
                    print(f"    furthest right x={reach} at rows {where[0]}..{where[-1]}")
            finally:
                shape.Delete()
        # Say which preset was actually drawn, read from the file itself.
        check = sheet.Shapes.AddShape(RIGHT_BRACE, LEFT_PT, TOP_PT, 40.0, 200.0)
        named = SCRATCH / "named.xlsx"
        if named.exists():
            named.unlink()
        book.SaveAs(str(named), FileFormat=51)
        check.Delete()
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    import re
    import zipfile

    holds = zipfile.ZipFile(SCRATCH / "named.xlsx")
    for part in holds.namelist():
        if part.startswith("xl/drawings/") and part.endswith(".xml"):
            found = re.findall(r'prst="([^"]+)"', holds.read(part).decode("utf-8", "replace"))
            print(f"  the preset drawn was: {found}")
    (SCRATCH / "_outline.json").write_text(json.dumps(held_all), encoding="utf-8")
    print(f"outlines in {SCRATCH / '_outline.json'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
