# -*- coding: utf-8 -*-
"""What does a Latin-faced cell ask a row for, and what is the Japanese face
it is paired with worth?

`fies_t2` — the lowest-scoring workbook — puts Times New Roman and Century in a
sheet whose Normal style is `Terminal 14`, and Excel gives those rows three
pixels more than either face's own line. A Japanese cell carries TWO faces (the
「日本語用」 and the 「英数字用」 of the font dialog) and xlsx has a slot for only
one, so a `<name val="Century"/>` cell keeps whatever Japanese face it inherits
— and the row's height is settled with both of them in hand.

This dresses a row in the Japanese face, sets ONLY the Latin one on the cell
(which is what `Font.Name` does with a Latin name in Japanese Excel), and reads
the row back. Two controls stand beside every arm: the Latin face alone, in a
row dressed one point, and the Japanese face alone.

    python tools\\metrics\\_xlsx_row_companion.py
"""

from __future__ import annotations

import sys

import win32com.client

WORDS = "Notes : 1. Change over the year"
COMPANIONS = [("Terminal", 14.0), ("ＭＳ 明朝", 10.0), ("ＭＳ 明朝", 11.0),
              ("ＭＳ Ｐゴシック", 11.0), ("游ゴシック", 11.0), ("メイリオ", 11.0),
              ("ＭＳ 明朝", 14.0)]
LATINS = [("Century", 9.0), ("Century", 11.0), ("Times New Roman", 10.0),
          ("Times New Roman", 11.0), ("Arial", 10.0), ("Calibri", 11.0)]


def main() -> int:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        at = [1]

        def ask(dress: tuple[str, float] | None, face: tuple[str, float] | None) -> int:
            row = at[0]
            at[0] += 1
            if dress is not None:
                sheet.Rows(row).Font.Name = dress[0]
                sheet.Rows(row).Font.Size = dress[1]
            else:
                sheet.Rows(row).Font.Name = "Calibri"
                sheet.Rows(row).Font.Size = 1.0
            cell = sheet.Cells(row, 2)
            cell.Value = WORDS
            if face is not None:
                cell.Font.Name = face[0]
                cell.Font.Size = face[1]
            sheet.Rows(row).AutoFit()
            return round(sheet.Rows(row).RowHeight * 96 / 72)

        alone = {latin: ask(None, latin) for latin in LATINS}
        print("  the Latin face alone, in a row dressed one point:")
        for latin, px in alone.items():
            print(f"    {latin[0]:<18}{latin[1]:>5}  {px:>3}px")
        print()
        print("  companion            own   " + "".join(
            f"{name[:7]:>8}{size:<3.0f}" for name, size in LATINS))
        for dress in COMPANIONS:
            own = ask(dress, None)
            row = []
            for latin in LATINS:
                got = ask(dress, latin)
                row.append(f"{got:>4}({got - max(alone[latin], own):+d})")
            print(f"  {dress[0]:<14}{dress[1]:>5}  {own:>4}   " + " ".join(row))
        print()
        print("  each cell: the row's height, and (its distance from"
              " max(Latin alone, companion alone))")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
