# -*- coding: utf-8 -*-
"""What Excel does when a line would start with a character that may not.

The advance is settled — a fullwidth glyph spends one em (SX18) — and what
is left is the break. A line holds `capacity` glyphs; put a 、 at position
capacity + 1 and Excel must do one of two things: hang it past the edge so
the line holds one more, or push its neighbour down so the line holds one
fewer. Excel's own PDF says which, because it writes down where every glyph
went.
"""
import json
import os
import sys
import zipfile

import fitz
import win32com.client

OUT_DIR = r"pipeline_data\repros\kinsoku"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="2"/><charset val="128"/></font><font><sz val="12"/><name val="游ゴシック"/><family val="2"/><charset val="128"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"""

SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A2"/><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="40" customWidth="1"/></cols><sheetData><row r="1" ht="400" customHeight="1"><c r="A1" s="1" t="inlineStr"><is><t>{text}</t></is></c></row></sheetData></worksheet>"""


def build(path, text):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/styles.xml", STYLES)
        z.writestr("xl/worksheets/sheet1.xml", SHEET.format(text=text))


def pdf_lines(pdf_path):
    document = fitz.open(pdf_path)
    page = document[0]
    lines = []
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            glyphs = [(char["c"], char["origin"][0], char["origin"][1])
                      for span in line["spans"] for char in span["chars"]]
            if glyphs:
                lines.append(glyphs)
    document.close()
    lines.sort(key=lambda glyphs: (round(glyphs[0][2], 1), glyphs[0][1]))
    return lines


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    root = os.path.abspath(OUT_DIR)
    # The plain case fixes the capacity; the rest place a character that may
    # not start a line exactly one past it.
    plans = [
        ("plain", "あ" * 60),
        ("touten", "あ" * 18 + "、" + "あ" * 30),
        ("kuten", "あ" * 18 + "。" + "あ" * 30),
        ("close", "あ" * 18 + "」" + "あ" * 30),
        ("small", "あ" * 18 + "っ" + "あ" * 30),
        ("open", "あ" * 17 + "「" + "あ" * 30),
        ("two", "あ" * 17 + "、、" + "あ" * 30),
    ]
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    out = []
    try:
        for tag, text in plans:
            name = "k_%s.xlsx" % tag
            build(os.path.join(root, name), text)
            wb = excel.Workbooks.Open(os.path.join(root, name), 0, False)
            pdf = os.path.join(root, "k_%s.pdf" % tag)
            if os.path.exists(pdf):
                os.remove(pdf)
            wb.ExportAsFixedFormat(0, pdf)
            wb.Close(False)
            lines = pdf_lines(pdf)
            print("== %s" % tag)
            entry = {"case": tag, "lines": []}
            for number, glyphs in enumerate(lines[:4]):
                shown = "".join(g[0] for g in glyphs)
                print("   line %d: %2d glyphs  %s" % (
                    number, len(glyphs),
                    shown[:12] + ("…" + shown[-4:] if len(shown) > 16 else "")))
                entry["lines"].append({"glyphs": len(glyphs), "text": shown})
            out.append(entry)
    finally:
        excel.Quit()
    with open(r"pipeline_data\com_measurements\xlsx_kinsoku.json", "w",
              encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
