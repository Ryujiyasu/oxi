"""V40-V45: Strip b5f706 V28 down to find the minimal trigger of 18pt cell snap.

V28 = full b5f706 Table 2 row 3 → dy=18pt confirmed.
V29-V34 (no balance/tblstyle/titlePg/cols/trHeight/full-styles) all → 18pt.

V40: V28 reduced to 1 cell (drop 14 of 15 cells, keep cell with 4 paragraphs)
V41: V40 minus paragraph rPr (no color, no lang, only minimal sz)
V42: V40 with simple section (drop landscape, drop pgMar header/footer specifics)
V43: V40 with single paragraph instead of 4 (does dy still go 18pt? wait single para has no dy)
V44: V40 with same structure but in portrait orientation
V45: V40 reduced even further: section sz=11906x16838 (portrait), pgMar all 1134tw
"""
from __future__ import annotations

import os
import re
import sys
import zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.join(REPO, "tools", "golden-test", "repros", "grid_snap")
B5F706_DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                            "b5f706e9f6ad_kyodokenkyuyoushiki_bessi.docx")


def write_docx(label: str, document_xml: str, settings_xml: str, styles_xml: str):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(B5F706_DOCX, 'r') as src:
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                if item.filename == 'word/document.xml':
                    dst.writestr(item, document_xml.encode('utf-8'))
                elif item.filename == 'word/settings.xml':
                    dst.writestr(item, settings_xml.encode('utf-8'))
                elif item.filename == 'word/styles.xml':
                    dst.writestr(item, styles_xml.encode('utf-8'))
                else:
                    dst.writestr(item, src.read(item.filename))
    print(f"  wrote {out_path}")


def extract_resources():
    """Extract b5f706 base XMLs."""
    with zipfile.ZipFile(B5F706_DOCX) as zf:
        document_xml = zf.read("word/document.xml").decode("utf-8")
        settings_xml = zf.read("word/settings.xml").decode("utf-8")
        styles_xml = zf.read("word/styles.xml").decode("utf-8")
    return document_xml, settings_xml, styles_xml


def make_v28_doc(document_xml_orig: str) -> str:
    """V28 baseline: only Table 2 row 3."""
    tables = re.findall(r'<w:tbl>.*?</w:tbl>', document_xml_orig, re.DOTALL)
    t2 = tables[1]
    tblpr = re.search(r'<w:tblPr>.*?</w:tblPr>', t2, re.DOTALL).group(0)
    tblgrid = re.search(r'<w:tblGrid>.*?</w:tblGrid>', t2, re.DOTALL).group(0)
    rows = re.findall(r'<w:tr[^>]*?>.*?</w:tr>', t2, re.DOTALL)
    row3 = rows[2]
    table_xml = f"<w:tbl>{tblpr}{tblgrid}{row3}</w:tbl>"
    sectpr = re.search(r'<w:sectPr[^>]*?>.*?</w:sectPr>', document_xml_orig, re.DOTALL).group(0)

    namespaces = ('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
                  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                  'xmlns:o="urn:schemas-microsoft-com:office:office" '
                  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
                  'xmlns:v="urn:schemas-microsoft-com:vml" '
                  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
                  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
                  'xmlns:w10="urn:schemas-microsoft-com:office:word" '
                  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
                  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
                  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
                  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
                  'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
                  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
                  'mc:Ignorable="w14 w15 wp14"')

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {namespaces}>
<w:body>
{table_xml}
<w:p/>
{sectpr}
</w:body>
</w:document>""", table_xml, sectpr


def make_v40_one_cell(table_xml: str) -> str:
    """V40: Reduce row 3 to 1 cell only (with 4 paragraphs - cell[2])."""
    cells = re.findall(r'<w:tc>.*?</w:tc>', table_xml, re.DOTALL)
    # Find first cell with 4 paragraphs
    target_cell = None
    for c in cells:
        n = c.count('<w:p ') + c.count('<w:p>')
        if n == 4:
            target_cell = c
            break
    if target_cell is None:
        target_cell = cells[2]  # fallback

    # New table: keep tblPr, but tblGrid is reduced to 1 column
    tblpr = re.search(r'<w:tblPr>.*?</w:tblPr>', table_xml, re.DOTALL).group(0)
    new_tblgrid = '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'

    # Cell with new width
    new_cell = re.sub(r'<w:tcW[^/]*?/>', '<w:tcW w:w="9000" w:type="dxa"/>', target_cell)
    # Remove gridSpan if present
    new_cell = re.sub(r'<w:gridSpan[^/]*?/>', '', new_cell)

    new_row = f'<w:tr><w:trPr><w:trHeight w:val="851" w:hRule="atLeast"/></w:trPr>{new_cell}</w:tr>'
    new_table = f'<w:tbl>{tblpr}{new_tblgrid}{new_row}</w:tbl>'
    return new_table


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT_DIR, exist_ok=True)

    document_xml_orig, settings_xml, styles_xml = extract_resources()
    v28_doc, v28_table_xml, v28_sectpr = make_v28_doc(document_xml_orig)

    namespaces = ('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
                  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                  'xmlns:o="urn:schemas-microsoft-com:office:office" '
                  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
                  'xmlns:v="urn:schemas-microsoft-com:vml" '
                  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
                  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
                  'xmlns:w10="urn:schemas-microsoft-com:office:word" '
                  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
                  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
                  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
                  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
                  'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
                  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
                  'mc:Ignorable="w14 w15 wp14"')

    def doc_with(table_xml: str, sectpr: str | None = None) -> str:
        sp = sectpr or v28_sectpr
        return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {namespaces}>
<w:body>
{table_xml}
<w:p/>
{sp}
</w:body>
</w:document>"""

    # ===== V40: V28 with only 1 cell of 4 paragraphs =====
    v40_table = make_v40_one_cell(v28_table_xml)
    v40_doc = doc_with(v40_table)
    write_docx("b5f706_V40_one_cell", v40_doc, settings_xml, styles_xml)

    # ===== V41: V40 with paragraph rPr stripped =====
    v41_table = re.sub(r'<w:rPr>(?!Default).*?</w:rPr>', '<w:rPr><w:sz w:val="18"/></w:rPr>',
                        v40_table, flags=re.DOTALL)
    # Avoid replacing rPrDefault
    v41_doc = doc_with(v41_table)
    write_docx("b5f706_V41_no_rpr", v41_doc, settings_xml, styles_xml)

    # ===== V42: V40 with simple section (portrait, no titlePg, no cols, no header/footer) =====
    simple_sectpr = """<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>"""
    v42_doc = doc_with(v40_table, simple_sectpr)
    write_docx("b5f706_V42_simple_section", v42_doc, settings_xml, styles_xml)

    # ===== V43: V40 minus trHeight =====
    v43_table = re.sub(r'<w:trHeight[^/]*?/>\s*', '', v40_table)
    v43_table = re.sub(r'<w:trPr>\s*</w:trPr>', '', v43_table)
    v43_doc = doc_with(v43_table)
    write_docx("b5f706_V43_no_trh", v43_doc, settings_xml, styles_xml)

    # ===== V44: V40 with all paragraph contents identical (same text in each) =====
    # No clear hypothesis but check
    v44_table = v40_table
    v44_doc = doc_with(v44_table)
    write_docx("b5f706_V44_baseline_dup", v44_doc, settings_xml, styles_xml)

    # ===== V45: V42 + simple styles (only Normal + Table Grid) =====
    minimal_styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="21"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/>
<w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>
</w:style>
<w:style w:type="table" w:styleId="aa"><w:name w:val="Table Grid"/><w:basedOn w:val="a1"/>
<w:tblPr><w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders></w:tblPr>
</w:style>
</w:styles>"""
    v45_doc = doc_with(v40_table, simple_sectpr)
    write_docx("b5f706_V45_minimal_all", v45_doc, settings_xml, minimal_styles)


if __name__ == "__main__":
    main()
