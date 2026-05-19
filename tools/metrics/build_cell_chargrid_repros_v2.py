"""V2: add vMerge + table style + other b35123 features to find the bug trigger.

V1 repros gave Word=6 chars (matching Oxi). b35123 gives Word=4 chars. The difference
must be one of the b35123-specific table properties.
"""
import os, zipfile

OUT_DIR = os.path.abspath("tools/metrics/cell_chargrid_repro_v2")
os.makedirs(OUT_DIR, exist_ok=True)

CT_DOC_STYLES = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'

STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="af"><w:name w:val="Table Grid"/><w:basedOn w:val="a1"/><w:uiPriority w:val="39"/><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/></w:rPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>'''

RPR = '<w:rFonts w:asciiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia" w:hint="eastAsia"/><w:sz w:val="21"/><w:szCs w:val="21"/>'


def build_doc(label: str, tc0_extras: str, with_continue_row: bool):
    """Build doc with cell0 specific extras (vMerge etc)."""
    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1134" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/></w:sectPr>'
    # Build the cell row
    text = "組織的管理措置"
    p_cell0 = f'<w:p><w:pPr><w:rPr>{RPR}</w:rPr></w:pPr><w:r><w:rPr>{RPR}</w:rPr><w:t>{text}</w:t></w:r></w:p>'
    p_cell1 = f'<w:p><w:pPr><w:rPr>{RPR}</w:rPr></w:pPr><w:r><w:rPr>{RPR}</w:rPr><w:t>□ 匿名データの適正管理に係る基本方針を定めること。</w:t></w:r></w:p>'
    row1 = (
        f'<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="1271" w:type="dxa"/>{tc0_extras}</w:tcPr>{p_cell0}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="7796" w:type="dxa"/></w:tcPr>{p_cell1}</w:tc>'
        f'</w:tr>'
    )
    # Optional continuation row (matches the multi-row form in b35123)
    if with_continue_row:
        p_cell0_b = f'<w:p><w:pPr><w:rPr>{RPR}</w:rPr></w:pPr></w:p>'  # empty for vMerge continue
        p_cell1_b = f'<w:p><w:pPr><w:rPr>{RPR}</w:rPr></w:pPr><w:r><w:rPr>{RPR}</w:rPr><w:t>□ 適正管理に関する考え方等を盛り込んだ内容</w:t></w:r></w:p>'
        # vMerge continue (no val) on continuation cell0
        row2 = (
            f'<w:tr>'
            f'<w:tc><w:tcPr><w:tcW w:w="1271" w:type="dxa"/><w:vMerge/></w:tcPr>{p_cell0_b}</w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="7796" w:type="dxa"/></w:tcPr>{p_cell1_b}</w:tc>'
            f'</w:tr>'
        )
    else:
        row2 = ''
    table = (
        f'<w:tbl>'
        f'<w:tblPr><w:tblStyle w:val="af"/><w:tblW w:w="9067" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="1197"/><w:gridCol w:w="7870"/></w:tblGrid>'
        f'{row1}{row2}'
        f'</w:tbl>'
    )
    doc = (f'<?xml version="1.0"?>'
           f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{table}{sect}</w:body></w:document>')
    path = os.path.join(OUT_DIR, f"{label}.docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT_DOC_STYLES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/document.xml", doc)
    print(f"Built {path}")


# Variants
build_doc("V_baseline", "", False)  # plain (= v1 reproduction, expect 6 chars)
build_doc("V_vMergeR", '<w:vMerge w:val="restart"/>', False)  # add vMerge restart only
build_doc("V_vMergeR_2rows", '<w:vMerge w:val="restart"/>', True)  # vMerge restart + continuation
build_doc("V_2rows_noMerge", "", True)  # 2 rows, no vMerge

print("Done.")
