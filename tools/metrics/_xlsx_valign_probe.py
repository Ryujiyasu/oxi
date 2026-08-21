# -*- coding: utf-8 -*-
"""Where does Excel put a line of text inside its row, top, middle or bottom?

The row's height is now exact, so what is left is where the letters sit in it.
Excel is asked directly: a sheet of cells that vary only in font, size, row
height and vertical alignment is exported to PDF, and the PDF says where each
baseline landed and where each cell's border ran. The difference between the
two, in pixels, is the law.

    python tools\\metrics\\_xlsx_valign_probe.py          # build, export, read
"""
import subprocess
import sys
from pathlib import Path

SCRATCH = Path(r"C:\tmp\xlsx_valign")
BOOK = SCRATCH / "valign_probe.xlsx"
PDF = SCRATCH / "valign_probe.pdf"

FONTS = [("ＭＳ ゴシック", 10.0), ("ＭＳ ゴシック", 11.0), ("ＭＳ 明朝", 11.0),
         ("Calibri", 11.0), ("ＭＳ Ｐゴシック", 9.0), ("Meiryo UI", 12.0)]
HEIGHTS = [None, 20.0, 30.0, 45.0]        # points; None = leave it natural
PLACES = ["top", "center", "bottom"]


def build():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, Side

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.title = "probe"
    edge = Side(style="thin", color="FF0000")
    frame = Border(left=edge, right=edge, top=edge, bottom=edge)

    rows = []
    row = 1
    for face, points in FONTS:
        for height in HEIGHTS:
            for place in PLACES:
                cell = sheet.cell(row=row, column=1, value="あA")
                cell.font = Font(name=face, size=points)
                cell.alignment = Alignment(vertical=place, horizontal="left")
                cell.border = frame
                if height is not None:
                    sheet.row_dimensions[row].height = height
                sheet.column_dimensions["A"].width = 14
                rows.append((row, face, points, height, place))
                row += 1
    book.save(BOOK)
    return rows


def export():
    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(BOOK), 0, False)
        sheet = book.Worksheets(1)
        heights = {row: sheet.Rows(row).Height for row in range(1, sheet.UsedRange.Rows.Count + 1)}
        sheet.PageSetup.Zoom = 100
        book.ExportAsFixedFormat(0, str(PDF))
        book.Close(False)
    finally:
        excel.Quit()
    return heights


def read(rows, heights):
    """Pair each line of text with the two rules that bracket it."""
    import fitz

    document = fitz.open(PDF)
    lines = []                       # (page, baseline, top rule, bottom rule, glyph top)
    for number, page in enumerate(document):
        rules = set()
        for drawing in page.get_drawings():
            rect = drawing["rect"]
            if rect.height < 2 and rect.width > 20:
                rules.add(round((rect.y0 + rect.y1) / 2, 2))
        rules = sorted(rules)
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                spans = [span for span in line["spans"] if span["text"].strip()]
                if not spans:
                    continue
                baseline = spans[0]["origin"][1]
                above = [y for y in rules if y < baseline - 0.5]
                below = [y for y in rules if y > baseline - 0.5]
                if not above or not below:
                    continue
                lines.append((number, baseline, above[-1], below[0],
                              min(span["bbox"][1] for span in spans)))
    lines.sort(key=lambda item: (item[0], item[1]))
    print(f"{len(lines)} lines of text bracketed by rules, {len(rows)} probes")

    print()
    print(f"{'font':<16}{'pt':>5}{'row px':>8}{'place':>8}"
          f"{'top→base':>10}{'base→bot':>10}{'measured px':>13}")
    for index, (row, face, points, asked, place) in enumerate(rows):
        if index >= len(lines):
            break
        _, baseline, top, bottom, glyph = lines[index]
        row_px = heights.get(row, 0) / 0.75
        print(f"{face:<16}{points:>5.0f}{row_px:>8.0f}{place:>8}"
              f"{(baseline - top) / 0.75:>10.2f}{(bottom - baseline) / 0.75:>10.2f}"
              f"{(bottom - top) / 0.75:>13.2f}")


def main():
    rows = build()
    heights = export()
    read(rows, heights)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
