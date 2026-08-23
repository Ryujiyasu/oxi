"""Does a wrapping CELL hang a trailing 句読点 the way a shape does?

`_xlsx_shape_hang.py` found that a shape's line lets 「。」 and 「、」 hang their
whole em past the room, and nothing else. The breaker is shared between cells
and shapes, so before that rule is put in the shared path this asks the same
question of a cell.

The arm: one column, its width swept across the point where the last character
stops fitting, the cell wrapping, the last character either 「。」 or 「あ」. What
is read is how many lines the row grew to — a cell that hangs stays one line
where a cell that does not becomes two.

Run: python tools/metrics/_xlsx_cell_hang.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import win32com.client

FACE = "ＭＳ ゴシック"   # every glyph one em
SIZE = 12.0
BODY = "いろはにほへとちりぬるをわかよたれそ"   # 18 characters


def main() -> int:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Cells.Font.Name = FACE
        sheet.Cells.Font.Size = SIZE
        # One character of ＭＳ ゴシック 12pt is one em; Excel states a column in
        # units of the standard font's zero, so the width is swept in points
        # through the crossing instead of guessed.
        print(f"{FACE} {SIZE:.0f}pt, body {len(BODY)} characters + a last one")
        print("column width in points; 'lines' is the row height divided by one line")
        # The one-line height, measured on a text that cannot wrap — calibrating
        # it from the first arm is how this read 28.50 for everything and called
        # a two-line row "one line".
        probe = sheet.Range("A1")
        probe.EntireColumn.ColumnWidth = 40
        probe.Value = "あ"
        probe.WrapText = True
        probe.EntireRow.AutoFit()
        one = probe.Height
        print(f"  one line is {one:.2f}pt")
        for last in ("。", "、", "」", "あ"):
            print(f"  last character 「{last}」")
            for short_px in (0, 2, 8, 14, 16, 18, 20):
                want_px = (len(BODY) + 1) * SIZE * 96 / 72 - short_px
                cell = sheet.Range("A1")
                cell.EntireColumn.ColumnWidth = 40
                cell.Value = BODY + last
                cell.WrapText = True
                # Excel states a column in characters; walk it until the drawn
                # width matches the pixels asked for.
                lo, hi = 1.0, 90.0
                for _ in range(40):
                    mid = (lo + hi) / 2
                    cell.EntireColumn.ColumnWidth = mid
                    if cell.Width * 96 / 72 < want_px:
                        lo = mid
                    else:
                        hi = mid
                cell.EntireColumn.ColumnWidth = lo
                cell.EntireRow.AutoFit()
                high = cell.Height
                lines = round(high / one) if one else 0
                print(f"    room {cell.Width * 96 / 72:6.1f} (asked {want_px:6.1f})"
                      f"  height {high:5.2f}  lines {lines}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
