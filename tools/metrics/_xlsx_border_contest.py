"""When two cells state a border on the edge they share, which one is drawn?

`R6kessan` stacks a row whose cells say `bottom thin` on a row whose cells say
`top double`. Excel draws the double — two lines with a white gap — and
nothing in between. Oxi draws both, so the gap fills in and the double reads
as a solid three-pixel rule. Three workbooks in the corpus do this.

So Excel does not draw both. This asks which it keeps, by stacking two cells
and sweeping the style each states at the edge between them, then reading the
pixels off Excel's own picture.

What is reported per pair: the rows of ink at the boundary, and which of the
two styles drawn alone matches them.

Run: python tools/metrics/_xlsx_border_contest.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_border_contest")
# Excel's own names for the styles, as XlLineStyle/XlBorderWeight pairs.
STYLES = {
    "thin": (1, 2),          # xlContinuous, xlThin
    "medium": (1, -4138),    # xlContinuous, xlMedium
    "thick": (1, 4),         # xlContinuous, xlThick
    "double": (-4119, 4),    # xlDouble, xlThick
    "dashed": (-4115, 2),
    "dotted": (-4118, 2),
    "hair": (1, 1),          # xlContinuous, xlHairline
}
ORDER = ["thin", "medium", "thick", "double", "dashed", "dotted", "hair"]
TOP_PT, LEFT_PT = 40.0, 40.0


def picture(sheet):
    for _ in range(8):
        try:
            sheet.Activate()
            sheet.Range("A1:H20").CopyPicture(Appearance=1, Format=2)
        except Exception:
            time.sleep(0.6)
            continue
        time.sleep(0.4)
        held = ImageGrab.grabclipboard()
        if held is not None:
            return held
    return None


def ink_rows(image, top: int, left: int, right: int, span: int = 12) -> list[int]:
    grey = np.asarray(image.convert("L"))[top - span : top + span, left:right]
    lit = (grey < 140).sum(axis=1)
    wide = (right - left) // 2
    return [y - span for y, n in enumerate(lit) if n > wide]


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:H20").Interior.Color = 0xFFFFFF
        sheet.Rows("2:3").RowHeight = 30.0
        sheet.Columns("B").ColumnWidth = 12.0
        upper, lower = sheet.Range("B2"), sheet.Range("B3")
        top = round(lower.Top * 96 / 72)
        left = round(lower.Left * 96 / 72) + 4
        right = left + round(lower.Width * 96 / 72) - 8

        # What each style looks like on its own, to compare the contests with.
        alone: dict[str, list[int]] = {}
        for name in ORDER:
            style, weight = STYLES[name]
            for cell in (upper, lower):
                cell.Borders.LineStyle = -4142      # xlLineStyleNone
            lower.Borders(8).LineStyle = style      # xlEdgeTop
            lower.Borders(8).Weight = weight
            held = picture(sheet)
            alone[name] = ink_rows(held, top, left, right) if held else []
            print(f"  {name:<8} alone: {alone[name]}")

        print("\n  upper says / lower says -> what is drawn")
        for above in ORDER:
            for below in ORDER:
                if above == below:
                    continue
                for cell in (upper, lower):
                    cell.Borders.LineStyle = -4142
                style, weight = STYLES[above]
                upper.Borders(9).LineStyle = style   # xlEdgeBottom
                upper.Borders(9).Weight = weight
                style, weight = STYLES[below]
                lower.Borders(8).LineStyle = style
                lower.Borders(8).Weight = weight
                held = picture(sheet)
                if held is None:
                    continue
                seen = ink_rows(held, top, left, right)
                matches = [name for name in ORDER if alone[name] == seen]
                print(f"    {above:<8} / {below:<8} -> {str(seen):<18} "
                      f"{'= ' + ', '.join(matches) if matches else '(neither alone)'}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
