# -*- coding: utf-8 -*-
r"""Everything Excel will say about `fies_t2`'s G131, beside a scratch twin.

The scratch and the book disagree about where a merged, distributed label
sits, at the same face, size, row height, column widths and Normal style. So
rather than keep adding candidates one at a time, this asks Excel for every
property of both cells and prints only what differs.

    python tools\metrics\_xlsx_fies_props.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
BOOK = REPO / "tools" / "golden-test" / "documents" / "xlsx" / "d1eb05860dd5_fies_t2.xlsx"
WIDTHS = [1.5, 12.09765625, 3.5, 1.59765625, 3.5]

CELL = ["HorizontalAlignment", "VerticalAlignment", "WrapText", "ShrinkToFit",
        "IndentLevel", "AddIndent", "Orientation", "ReadingOrder", "MergeCells",
        "NumberFormat", "Style", "Text", "Value"]
FONT = ["Name", "Size", "Bold", "Italic", "Underline", "Strikethrough",
        "Subscript", "Superscript", "FontStyle"]


def look(cell, sheet) -> dict:
    held = {}
    for name in CELL:
        try:
            held[name] = str(cell.__getattr__(name))
        except Exception as trouble:
            held[name] = f"<{type(trouble).__name__}>"
    for name in FONT:
        try:
            held["Font." + name] = str(cell.Font.__getattr__(name))
        except Exception as trouble:
            held["Font." + name] = f"<{type(trouble).__name__}>"
    try:
        held["MergeArea"] = cell.MergeArea.Address(False, False)
        held["MergeArea.Width"] = str(round(cell.MergeArea.Width, 3))
        held["MergeArea.Height"] = str(round(cell.MergeArea.Height, 3))
    except Exception:
        pass
    held["RowHeight"] = str(cell.RowHeight)
    held["Normal font"] = (f"{sheet.Parent.Styles('Normal').Font.Name} "
                           f"{sheet.Parent.Styles('Normal').Font.Size}")
    held["StandardWidth"] = str(sheet.StandardWidth)
    held["StandardHeight"] = str(sheet.StandardHeight)
    return held


def main() -> int:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(BOOK), ReadOnly=True)
    theirs = look(book.Worksheets(1).Range("G131"), book.Worksheets(1))
    book.Close(SaveChanges=False)

    made = excel.Workbooks.Add()
    sheet = made.Worksheets(1)
    for at, width in enumerate(WIDTHS):
        sheet.Columns(7 + at).ColumnWidth = width
    cell = sheet.Range("G131")
    cell.Value = "女性用洋服"
    cell.Font.Name = "ＭＳ 明朝"
    cell.Font.Size = 12
    cell.HorizontalAlignment = -4117
    sheet.Range("G131:K131").Merge()
    sheet.Rows(131).RowHeight = 14.25
    time.sleep(0.4)
    ours = look(sheet.Range("G131"), sheet)
    made.Close(SaveChanges=False)
    excel.Quit()

    print("  G131: the book, then a scratch twin — only what differs")
    same = 0
    for name in theirs:
        if theirs[name] == ours.get(name):
            same += 1
            continue
        print(f"    {name:<20} book={theirs[name]!r:<28} scratch={ours.get(name)!r}")
    print(f"  {same} of {len(theirs)} properties agree")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
