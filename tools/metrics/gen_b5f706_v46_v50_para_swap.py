"""V46-V50: Find the precise feature differing between V20 (11.5pt natural)
and V40 (18pt snap).

V20: handwritten paragraphs in b5f706-mimic table → 11.5pt
V40: b5f706 extracted row 3 reduced to 1 cell → 18.0pt

V46: V40 with paragraphs replaced by V20-style handwritten ones
V47: V46 with styles replaced by V20-plain (Normal only)
V48: V47 with settings.xml replaced by V20-plain (no balance flag)
V49: V48 packaged as fresh docx (no theme.xml, no fontTable.xml, no numbering.xml)
V50: V49 with paragraph mark rPr matching V20's
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

PLAIN_SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

PLAIN_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

NAMESPACES = ('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
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


def write_docx_inplace(label: str, document_xml: str, settings_xml: str,
                        styles_xml: str):
    """Copy b5f706 docx wholesale, replacing key XMLs."""
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


def write_docx_minimal(label: str, document_xml: str, settings_xml: str,
                        styles_xml: str):
    """Build a minimal docx WITHOUT theme.xml/fontTable.xml/etc."""
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/settings.xml", settings_xml)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("word/document.xml", document_xml)
    print(f"  wrote {out_path} (minimal package)")


def make_simple_para(text: str, sz_hp: int = 18, jc: str = "center") -> str:
    return (f'<w:p><w:pPr><w:jc w:val="{jc}"/></w:pPr>'
            f'<w:r><w:rPr><w:sz w:val="{sz_hp}"/></w:rPr><w:t>{text}</w:t></w:r></w:p>')


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT_DIR, exist_ok=True)

    # b5f706 originals
    with zipfile.ZipFile(B5F706_DOCX) as zf:
        document_xml = zf.read("word/document.xml").decode("utf-8")
        b5f706_settings = zf.read("word/settings.xml").decode("utf-8")
        b5f706_styles = zf.read("word/styles.xml").decode("utf-8")

    # Extract V40-style table (Table 2 row 3 reduced to 1 cell with 4 paragraphs)
    tables = re.findall(r'<w:tbl>.*?</w:tbl>', document_xml, re.DOTALL)
    t2 = tables[1]
    tblpr = re.search(r'<w:tblPr>.*?</w:tblPr>', t2, re.DOTALL).group(0)
    rows = re.findall(r'<w:tr[^>]*?>.*?</w:tr>', t2, re.DOTALL)
    row3 = rows[2]
    cells = re.findall(r'<w:tc>.*?</w:tc>', row3, re.DOTALL)
    target_cell = next((c for c in cells if c.count('<w:p ') + c.count('<w:p>') == 4), cells[2])

    # V40 (1 cell with original 4 paragraphs)
    new_tblgrid = '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
    cell_v40 = re.sub(r'<w:tcW[^/]*?/>', '<w:tcW w:w="9000" w:type="dxa"/>', target_cell)
    cell_v40 = re.sub(r'<w:gridSpan[^/]*?/>', '', cell_v40)
    new_row_v40 = f'<w:tr><w:trPr><w:trHeight w:val="851" w:hRule="atLeast"/></w:trPr>{cell_v40}</w:tr>'
    table_v40 = f'<w:tbl>{tblpr}{new_tblgrid}{new_row_v40}</w:tbl>'

    sectpr_orig = re.search(r'<w:sectPr[^>]*?>.*?</w:sectPr>', document_xml, re.DOTALL).group(0)

    simple_sectpr = """<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>"""

    # Replace 4 cell paragraphs with V20-style simple paragraphs
    simple_paragraphs = "".join(make_simple_para(f"V46 L{i}") for i in range(1, 5))
    cell_v46 = re.sub(
        r'(<w:tc>\s*<w:tcPr>.*?</w:tcPr>)(.*?)(</w:tc>)',
        lambda m: m.group(1) + simple_paragraphs + m.group(3),
        cell_v40, count=1, flags=re.DOTALL,
    )
    new_row_v46 = f'<w:tr><w:trPr><w:trHeight w:val="851" w:hRule="atLeast"/></w:trPr>{cell_v46}</w:tr>'
    table_v46 = f'<w:tbl>{tblpr}{new_tblgrid}{new_row_v46}</w:tbl>'

    def doc_with(table_xml: str, sectpr: str) -> str:
        return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {NAMESPACES}>
<w:body>
{table_xml}
<w:p/>
{sectpr}
</w:body>
</w:document>"""

    # ===== V46: V40 with paragraphs replaced by simple ones (still uses b5f706 styles + sectPr) =====
    write_docx_inplace("b5f706_V46_simple_paras",
                       doc_with(table_v46, sectpr_orig),
                       b5f706_settings, b5f706_styles)

    # ===== V47: V46 with simple section =====
    write_docx_inplace("b5f706_V47_simple_section",
                       doc_with(table_v46, simple_sectpr),
                       b5f706_settings, b5f706_styles)

    # ===== V48: V47 with plain styles =====
    write_docx_inplace("b5f706_V48_plain_styles",
                       doc_with(table_v46, simple_sectpr),
                       b5f706_settings, PLAIN_STYLES)

    # ===== V49: V48 with plain settings (no balance flag) =====
    write_docx_inplace("b5f706_V49_plain_settings",
                       doc_with(table_v46, simple_sectpr),
                       PLAIN_SETTINGS, PLAIN_STYLES)

    # ===== V50: V49 packaged minimal (no theme.xml, fontTable.xml, etc.) =====
    write_docx_minimal("b5f706_V50_minimal_pkg",
                       doc_with(table_v46, simple_sectpr),
                       PLAIN_SETTINGS, PLAIN_STYLES)


if __name__ == "__main__":
    main()
