"""Minimal cell+charGrid repros for b35123 col-width investigation.

Bug: Word col1 cell renders 4 chars (12.45pt/char), Oxi renders 6 chars (9.84pt/char).
Same font sz=21 (10.5pt), same cell tcW=1271tw (63.55pt), default tcMar 108tw L/R.
Oxi formula: char_pitch = font_size + charSpace/4096 = 10.5 - 0.66 = 9.84pt (COM-verified for BODY).
Word table cell: char pitch = ~12.45pt = font_size + 1.95pt EXPANSION (formula unknown).

Repro grid:
- charSpace values: 0, -1000, -2000, -2714, -3000, -4000
- cell tcW: 1271tw (same as b35123)
- font size: sz=21 (10.5pt) MS Mincho
- text: "組織的管理措置" (7 fullwidth CJK chars) — see how many wrap onto line 1
- Also vary tcW: 1000, 1271, 1600, 2000tw to see if width affects char pitch
"""
import os, zipfile

OUT_DIR = os.path.abspath("tools/metrics/cell_chargrid_repro")
os.makedirs(OUT_DIR, exist_ok=True)

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

RPR = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="21"/><w:szCs w:val="21"/>'

TBL_BORDERS = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'


def make_doc(charspace: int, tcw_col1: int, text: str = "組織的管理措置"):
    sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1134" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="{charspace}"/></w:sectPr>'
    tcw_col2 = 9070 - tcw_col1  # body width = 9070tw
    tbl_w = tcw_col1 + tcw_col2
    table = (
        f'<w:tbl>'
        f'<w:tblPr><w:tblW w:w="{tbl_w}" w:type="dxa"/>{TBL_BORDERS}</w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{tcw_col1}"/><w:gridCol w:w="{tcw_col2}"/></w:tblGrid>'
        f'<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="{tcw_col1}" w:type="dxa"/></w:tcPr>'
        f'<w:p><w:pPr><w:rPr>{RPR}</w:rPr></w:pPr><w:r><w:rPr>{RPR}</w:rPr><w:t>{text}</w:t></w:r></w:p>'
        f'</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="{tcw_col2}" w:type="dxa"/></w:tcPr>'
        f'<w:p><w:pPr><w:rPr>{RPR}</w:rPr></w:pPr><w:r><w:rPr>{RPR}</w:rPr><w:t>col2</w:t></w:r></w:p>'
        f'</w:tc>'
        f'</w:tr></w:tbl>'
    )
    return (f'<?xml version="1.0"?>'
            f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{table}{sect}</w:body></w:document>')


def build(label: str, charspace: int, tcw_col1: int):
    doc = make_doc(charspace, tcw_col1)
    path = os.path.join(OUT_DIR, f"{label}.docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc)
    print(f"Built {path}")


# Vary charSpace, keep tcW=1271 (b35123 actual)
for cs in [0, -1000, -2000, -2714, -3000, -4000]:
    sign = 'p' if cs >= 0 else 'n'
    build(f"C_cs{sign}{abs(cs):05d}_tcw1271", cs, 1271)

# Vary tcW with b35123 charSpace
for tcw in [1000, 1271, 1500, 2000, 3000]:
    build(f"C_csn02714_tcw{tcw:04d}", -2714, tcw)

print("Done.")
