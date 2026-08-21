# -*- coding: utf-8 -*-
"""Read the advances and the wrap width out of Excel's own PDF.

Every engine on this machine measures text differently from Excel — GDI steps
游ゴシック's す a pixel short of the em, DirectWrite's GDI-compatible layout runs
a fraction under GDI, and `Columns.AutoFit` inflates by a tenth. Excel's PDF
export does not guess: it writes the position it chose for every glyph, so a
PDF reader gives back the advance Excel used and the exact character each line
broke at.

Builds one probe workbook per (font, size, column width), exports it, and
reports per line: the first glyph's x, the advance between glyphs, how many
characters landed on the line, and the ink the line covers.
"""
import json
import os
import sys
import zipfile

import fitz
import win32com.client

OUT_DIR = r"pipeline_data\repros\pdf_glyphs"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="2"/><charset val="128"/></font><font><sz val="{size}"/><name val="{font}"/><family val="2"/><charset val="128"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"""

# B1 carries a single glyph, so the distance between the two cells' first
# glyphs is column A's width as the PDF actually laid it out — the screen's
# idea of the column need not survive the trip to print.
SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="{width}" customWidth="1"/><col min="2" max="2" width="12" customWidth="1"/></cols><sheetData><row r="1" ht="400" customHeight="1"><c r="A1" s="2" t="inlineStr"><is><t>{text}</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>[</t></is></c></row></sheetData></worksheet>"""


def build(path, font, size, width, text):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/styles.xml", STYLES.format(font=font, size=size))
        z.writestr("xl/worksheets/sheet1.xml",
                   SHEET.format(width=width, text=text))


def ruled_box(pdf_path):
    """The rectangle A1's border draws, in pixels: its left and right edges
    are the column's own, so the padding either side of the text can be read
    off directly instead of inferred."""
    document = fitz.open(pdf_path)
    page = document[0]
    verticals = []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] == "l":
                start, end = item[1], item[2]
                if abs(start.x - end.x) < 0.01 and abs(start.y - end.y) > 2:
                    verticals.append(start.x * 96.0 / 72.0)
            elif item[0] == "re":
                rect = item[1]
                verticals.extend([rect.x0 * 96.0 / 72.0, rect.x1 * 96.0 / 72.0])
    document.close()
    verticals = sorted(set(round(x, 3) for x in verticals))
    return verticals


def glyph_lines(pdf_path):
    """Per drawn line: the glyphs Excel placed, with their x positions in
    points. PDF points are 1/72in; the pixels everything else is measured in
    are 1/96in, so a point is 4/3 of a pixel."""
    document = fitz.open(pdf_path)
    page = document[0]
    lines = []
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            glyphs = []
            for span in line["spans"]:
                for char in span["chars"]:
                    glyphs.append((char["c"], char["bbox"][0], char["bbox"][2],
                                   char["origin"][0]))
            if glyphs:
                lines.append(glyphs)
    document.close()
    lines.sort(key=lambda glyphs: glyphs[0][3])
    return lines


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    root = os.path.abspath(OUT_DIR)
    # A wide row so every wrapped line is drawn, and a column width that the
    # earlier batteries put right at the knife edge.
    plans = [
        ("游ゴシック", 12, 40.0, "あ" * 60, "the plain fullwidth case"),
        ("游ゴシック", 12, 40.0, "すすすすすすすすすすすすすすすすすすすす", "GDI steps す narrow"),
        ("游ゴシック", 12, 40.0, "しししししししししししししししししししし", "GDI steps し narrower"),
        ("游ゴシック", 11, 40.0, "あ" * 60, "the +5 font"),
        ("メイリオ", 11, 40.0, "あ" * 60, "the +7 font at the same ppem"),
        ("Arial", 12, 40.0, "n" * 80, "+7 latin"),
        ("Calibri", 12, 40.0, "n" * 80, "+5 latin"),
    ]
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    results = []
    try:
        for index, (font, size, width, text, why) in enumerate(plans):
            name = "g%02d.xlsx" % index
            build(os.path.join(root, name), font, size, width, text)
            wb = excel.Workbooks.Open(os.path.join(root, name), 0, False)
            ws = wb.Worksheets(1)
            column_px = ws.Columns(1).Width / 0.75
            pdf = os.path.join(root, "g%02d.pdf" % index)
            if os.path.exists(pdf):
                os.remove(pdf)
            wb.ExportAsFixedFormat(0, pdf)
            wb.Close(False)

            lines = glyph_lines(pdf)
            # The marker in B1 sits on the first line, past all of A1's text.
            marker = None
            for glyphs in lines:
                for glyph in glyphs:
                    if glyph[0] == "[":
                        marker = glyph[3] * 96.0 / 72.0
            body = [[g for g in glyphs if g[0] != "["] for glyphs in lines]
            lines = [glyphs for glyphs in body if glyphs]
            first_x = lines[0][0][3] * 96.0 / 72.0 if lines else 0.0
            drawn_column = (marker - first_x) if marker else None
            edges = ruled_box(pdf)
            print("\n== %s %s, column %.0fpx (drawn %s) — %s" % (
                font, size, column_px,
                "%.2fpx" % drawn_column if drawn_column else "?", why))
            entry = {"font": font, "size": size, "column_px": column_px,
                     "drawn_column_px": drawn_column, "lines": []}
            if len(edges) >= 2:
                left, right = edges[0], edges[-1]
                widest = max((len(g) for g in lines), default=0)
                advance = ((lines[0][1][3] - lines[0][0][3]) * 96.0 / 72.0
                           if lines and len(lines[0]) > 1 else 0.0)
                print("   ruled %.2f..%.2f = %.2fpx | text starts %+.2f from "
                      "the left | %d x %.2f = %.2f leaves %+.2f at the right"
                      % (left, right, right - left, first_x - left, widest,
                         advance, widest * advance,
                         right - (first_x + widest * advance)))
                entry.update({"ruled_left": left, "ruled_right": right,
                              "pad_left": first_x - left,
                              "pad_right": right - (first_x + widest * advance),
                              "advance": advance, "capacity": widest})
            for number, glyphs in enumerate(lines[:3]):
                origins = [g[3] for g in glyphs]
                steps = [round((b - a) * 96.0 / 72.0, 3)
                         for a, b in zip(origins, origins[1:])]
                left_px = origins[0] * 96.0 / 72.0
                # where the line's ink ends, and where the next glyph would
                right_px = glyphs[-1][2] * 96.0 / 72.0
                advance = steps[0] if steps else 0.0
                covered = len(glyphs) * advance
                print("   line %d: %2d glyphs, first x %.2fpx, advance %s, "
                      "ink to %.2fpx, %d x advance = %.2fpx" % (
                          number, len(glyphs), left_px,
                          sorted(set(steps))[:3], right_px,
                          len(glyphs), covered))
                entry["lines"].append({
                    "glyphs": len(glyphs), "left_px": left_px,
                    "advance": advance, "right_px": right_px,
                    "steps": sorted(set(steps))[:5],
                })
            results.append(entry)
    finally:
        excel.Quit()
    with open(r"pipeline_data\com_measurements\xlsx_pdf_glyphs.json", "w",
              encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
