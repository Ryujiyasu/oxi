# -*- coding: utf-8 -*-
"""How far inside its cell does Excel start the text, face by face?

Excel's PDF places the first glyph 3.36px inside the column for 游ゴシック 11,
3.52 for 游ゴシック 12 and 3.68 for メイリオ 11 — and those are the same faces
whose wrapped lines keep 5, 7 and 7 pixels. So the inset may be the
allowance's other half. One workbook, one font a row, one export.
"""
import json
import os
import sys
import zipfile

import fitz
import win32com.client

OUT_DIR = r"pipeline_data\repros\pdf_inset"

FACES = [
    ("ＭＳ 明朝", 9), ("ＭＳ 明朝", 11), ("ＭＳ 明朝", 12), ("ＭＳ 明朝", 14),
    ("ＭＳ Ｐゴシック", 11), ("ＭＳ Ｐゴシック", 14),
    ("游ゴシック", 9), ("游ゴシック", 11), ("游ゴシック", 12), ("游ゴシック", 16),
    ("Yu Gothic UI", 11), ("Yu Gothic UI", 12),
    ("メイリオ", 10), ("メイリオ", 11),
    ("Meiryo UI", 10), ("Meiryo UI", 11),
    ("Century", 9), ("Terminal", 14),
]

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"""


def styles():
    fonts = ['<font><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="2"/>'
             '<charset val="128"/></font>']
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    for index, (face, size) in enumerate(FACES, start=1):
        fonts.append('<font><sz val="%s"/><name val="%s"/><family val="2"/>'
                     '<charset val="128"/></font>' % (size, face))
        xfs.append('<xf numFmtId="0" fontId="%d" fillId="0" borderId="1" '
                   'xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">'
                   '<alignment vertical="top"/></xf>' % index)
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="%d">%s</fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="%d">%s</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"""
            % (len(fonts), "".join(fonts), len(xfs), "".join(xfs)))


def sheet_xml():
    rows = []
    for index, (face, size) in enumerate(FACES, start=1):
        rows.append('<row r="%d" ht="42" customHeight="1"><c r="A%d" s="%d" '
                    't="inlineStr"><is><t>あああ</t></is></c></row>'
                    % (index, index, index))
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A%d"/><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="30" customWidth="1"/></cols><sheetData>%s</sheetData></worksheet>"""
            % (len(FACES), "".join(rows)))


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    root = os.path.abspath(OUT_DIR)
    path = os.path.join(root, "inset.xlsx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/styles.xml", styles())
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml())

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    pdf = os.path.join(root, "inset.pdf")
    try:
        wb = excel.Workbooks.Open(path, 0, False)
        if os.path.exists(pdf):
            os.remove(pdf)
        wb.ExportAsFixedFormat(0, pdf)
        wb.Close(False)
    finally:
        excel.Quit()

    document = fitz.open(pdf)
    page = document[0]
    # The column's left edge: the leftmost vertical rule on the page.
    edges = []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] == "l" and abs(item[1].x - item[2].x) < 0.01:
                edges.append(item[1].x * 96.0 / 72.0)
            elif item[0] == "re":
                edges.append(item[1].x0 * 96.0 / 72.0)
    left = min(edges) if edges else 0.0
    glyphs = []
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                for char in span["chars"]:
                    glyphs.append((char["origin"][1], char["origin"][0]))
    document.close()
    # One line a font, top to bottom, in the order the rows were written.
    firsts = {}
    for y, x in sorted(glyphs):
        firsts.setdefault(round(y, 1), x * 96.0 / 72.0)
    order = [firsts[y] for y in sorted(firsts)]
    print("column's left edge %.2fpx" % left)
    out = []
    for (face, size), x in zip(FACES, order):
        inset = x - left
        out.append({"face": face, "size": size, "inset_px": inset,
                    "inset_units": inset / 0.16})
        print("   %-16s %-4s first glyph %.2fpx -> inset %.2fpx (%.0f units)"
              % (face, size, x, inset, inset / 0.16))
    with open(r"pipeline_data\com_measurements\xlsx_pdf_inset.json", "w",
              encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
