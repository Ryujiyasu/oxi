# -*- coding: utf-8 -*-
"""What a Latin cell asks a row for, with the Japanese face held small.

`_xlsx_row_companion.py` dressed the whole row in the Japanese face, so the
BLANK cells asked for that face's own line and the row never fell below it —
which hid what the Latin cell itself was asking. This holds the blanks at ＭＳ
明朝 8 (14 pixels) and moves the companion through the workbook's NORMAL style,
where a Latin-named cell inherits it. Now the reading is the cell's own ask.

`fies_t2` says the ask is neither face's line but a line built from both:

    ask = max(baseline, companion baseline - 1) + max(descent, companion descent)

with the parts taken from `row_defaults` (row height, and how far down it the
baseline sits). This sweeps a second and third companion to see whether that
holds, or whether Terminal 14 was a special case.

    python tools\\metrics\\_xlsx_row_companion2.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
TABLE = REPO / "tools" / "oxi-xlsx-renderer" / "src" / "row_defaults.rs"
COMPANIONS = [("Terminal", 14.0), ("ＭＳ 明朝", 14.0), ("游ゴシック", 11.0),
              ("メイリオ", 11.0), ("ＭＳ Ｐゴシック", 12.0)]
LATINS = [("Century", 9.0), ("Century", 11.0), ("Century", 14.0),
          ("Times New Roman", 10.0), ("Times New Roman", 11.0), ("Arial", 10.0)]


def table() -> dict[tuple[str, int], tuple[int, int]]:
    held = {}
    for line in TABLE.read_text(encoding="utf-8").splitlines():
        found = re.match(r'\s*\("([^"]+)",\s*(\d+),\s*(\d+),\s*(\d+)\)', line)
        if found:
            held[(found.group(1), int(found.group(2)))] = (
                int(found.group(3)), int(found.group(4)))
    return held


def main() -> int:
    known = table()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    print("  blanks held at ＭＳ 明朝 8 (14px); the companion is the Normal style")
    print("  companion          latin              ask   said   own    (said = the model)")
    for companion in COMPANIONS:
        book = excel.Workbooks.Add()
        try:
            sheet = book.Worksheets(1)
            book.Styles("Normal").Font.Name = companion[0]
            book.Styles("Normal").Font.Size = companion[1]
            for column in range(1, 40):
                sheet.Columns(column).Font.Name = "ＭＳ 明朝"
                sheet.Columns(column).Font.Size = 8.0
            at = 2
            for latin in LATINS:
                cell = sheet.Cells(at, 2)
                cell.Value = "Notes 1 Change"
                cell.Font.Name = latin[0]
                cell.Font.Size = latin[1]
                sheet.Rows(at).AutoFit()
                ask = round(sheet.Rows(at).RowHeight * 96 / 72)
                own = known.get((latin[0], round(latin[1] * 4)))
                mate = known.get((companion[0], round(companion[1] * 4)))
                if own and mate:
                    base = max(own[1], mate[1] - 1)
                    down = max(own[0] - own[1], mate[0] - mate[1])
                    said = base + down
                else:
                    said = None
                flag = "" if said == ask else "  <<"
                print(f"  {companion[0]:<12}{companion[1]:>5}  {latin[0]:<16}{latin[1]:>5}"
                      f"  {ask:>4}  {str(said):>5}  {str(own):>9}{flag}")
                at += 1
        finally:
            book.Close(SaveChanges=False)
    excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
