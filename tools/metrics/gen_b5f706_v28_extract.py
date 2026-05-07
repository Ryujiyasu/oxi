"""V28: Extract b5f706 Table 2 row 3 verbatim, embed in minimal document.

Bottom-up approach: take exactly the cell that produces dy=18pt in b5f706
(Table 2 row 3 cell[2] with 4 paragraphs, fs=9pt, jc=center, vAlign=center,
trHeight=851 atLeast) and place it in an otherwise-empty docx with the same
section properties as b5f706.

If V28 reproduces dy=18pt, then iteratively remove features (V29, V30, ...)
until dy flips back to 13pt. The feature whose removal flips dy is the trigger.
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


def extract_table2_row3_xml(b5f706_path: str) -> tuple[str, str, str, str]:
    """Returns (table2_row3_full_table_xml, sectPr_xml, settings_xml, styles_xml)
    where table xml has only row 3 (we keep tblPr + tblGrid from original)."""
    with zipfile.ZipFile(b5f706_path) as zf:
        document_xml = zf.read("word/document.xml").decode("utf-8")
        settings_xml = zf.read("word/settings.xml").decode("utf-8")
        styles_xml = zf.read("word/styles.xml").decode("utf-8")

    # Find tables
    tables = re.findall(r'<w:tbl>.*?</w:tbl>', document_xml, re.DOTALL)
    t2 = tables[1]  # Table 2

    # Extract tblPr, tblGrid
    tblpr = re.search(r'<w:tblPr>.*?</w:tblPr>', t2, re.DOTALL).group(0)
    tblgrid = re.search(r'<w:tblGrid>.*?</w:tblGrid>', t2, re.DOTALL).group(0)
    rows = re.findall(r'<w:tr[^>]*?>.*?</w:tr>', t2, re.DOTALL)
    row3 = rows[2]  # row index 2 = Table 2's 3rd row

    table_xml = f"<w:tbl>{tblpr}{tblgrid}{row3}</w:tbl>"

    # Section properties: keep b5f706's exact section
    sectpr = re.search(r'<w:sectPr[^>]*?>.*?</w:sectPr>', document_xml, re.DOTALL).group(0)

    return table_xml, sectpr, settings_xml, styles_xml


def write_docx_replace(label: str, document_xml: str, settings_xml: str, styles_xml: str):
    """Copy b5f706 docx wholesale, replacing document.xml/settings.xml/styles.xml."""
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


# Backwards-compatible alias
write_docx = write_docx_replace


def build_doc(table_xml: str, sectpr_xml: str) -> str:
    # Use the same namespace declarations as b5f706 (extracted XML carries
    # w14:paraId, mc:Ignorable, etc. so we need the full xmlns set).
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">
<w:body>
{table_xml}
<w:p/>
{sectpr_xml}
</w:body>
</w:document>"""


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT_DIR, exist_ok=True)

    table_xml, sectpr, settings_xml, styles_xml = extract_table2_row3_xml(B5F706_DOCX)

    # V28: full extract (Table 2 row 3 only, all features intact)
    document_xml = build_doc(table_xml, sectpr)
    write_docx("b5f706_V28_extract", document_xml, settings_xml, styles_xml)

    # V29: V28 minus balanceSingleByteDoubleByteWidth flag
    settings_no_balance = re.sub(r'<w:balanceSingleByteDoubleByteWidth/>\s*', '', settings_xml)
    write_docx("b5f706_V29_no_balance", document_xml, settings_no_balance, styles_xml)

    # V30: V28 minus tblStyle reference
    table_no_tblstyle = re.sub(r'<w:tblStyle[^/]*?/>\s*', '', table_xml)
    document_v30 = build_doc(table_no_tblstyle, sectpr)
    write_docx("b5f706_V30_no_tblstyle", document_v30, settings_xml, styles_xml)

    # V31: V28 minus titlePg in section
    sectpr_no_titlepg = re.sub(r'<w:titlePg/>\s*', '', sectpr)
    document_v31 = build_doc(table_xml, sectpr_no_titlepg)
    write_docx("b5f706_V31_no_titlepg", document_v31, settings_xml, styles_xml)

    # V32: V28 minus cols (single column section)
    sectpr_no_cols = re.sub(r'<w:cols[^/]*?/>\s*', '', sectpr)
    document_v32 = build_doc(table_xml, sectpr_no_cols)
    write_docx("b5f706_V32_no_cols", document_v32, settings_xml, styles_xml)

    # V33: V28 minus row 3 trHeight constraint (no trHeight)
    table_no_trh = re.sub(r'<w:trHeight[^/]*?/>\s*', '', table_xml)
    document_v33 = build_doc(table_no_trh, sectpr)
    write_docx("b5f706_V33_no_trheight", document_v33, settings_xml, styles_xml)

    # V34: V28 with simplified styles (drop everything except table grid + Normal)
    # Provide a minimal styles.xml that still defines style "aa" identically
    # but drops all unused other styles
    minimal_styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
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
    write_docx("b5f706_V34_minimal_styles", document_xml, settings_xml, minimal_styles)


if __name__ == "__main__":
    main()
