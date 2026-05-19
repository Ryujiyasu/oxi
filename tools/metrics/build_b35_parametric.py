"""Parametric repro for S110: vary fs × charSpace × tcW × content to triangulate
Word's char_pitch formula under balanceSingleByteDoubleByteWidth + 4-row vMerge.
"""
import os, zipfile

OUT_DIR = os.path.abspath("tools/metrics/b35_parametric_repro")
os.makedirs(OUT_DIR, exist_ok=True)

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:balanceSingleByteDoubleByteWidth/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>'''


def make_doc(sz_halfpt: int, charspace: int, tcw_col1: int, text: str = "組織的管理措置"):
    """sz_halfpt: 18=9pt, 21=10.5pt, 24=12pt, 28=14pt"""
    rpr = f'<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="{sz_halfpt}"/><w:szCs w:val="{sz_halfpt}"/>'
    sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1134" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="{charspace}"/></w:sectPr>'
    tcw_col2 = 9070 - tcw_col1
    borders = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'

    # 5 rows: row0 = header, row1 = restart cell0 with text, rows 2-4 = vMerge continuation
    def cell0_restart():
        return (f'<w:tc><w:tcPr><w:tcW w:w="{tcw_col1}" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>'
                f'<w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>{text}</w:t></w:r></w:p>'
                f'</w:tc>')

    def cell0_continue():
        return (f'<w:tc><w:tcPr><w:tcW w:w="{tcw_col1}" w:type="dxa"/><w:vMerge/></w:tcPr>'
                f'<w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr></w:p>'
                f'</w:tc>')

    def cell1(label: str):
        return (f'<w:tc><w:tcPr><w:tcW w:w="{tcw_col2}" w:type="dxa"/></w:tcPr>'
                f'<w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>{label}</w:t></w:r></w:p>'
                f'</w:tc>')

    # row0 (header — no vMerge involvement; just to mimic real table)
    row0 = (f'<w:tr>'
            f'<w:tc><w:tcPr><w:tcW w:w="{tcw_col1}" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>区分</w:t></w:r></w:p></w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="{tcw_col2}" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>内容</w:t></w:r></w:p></w:tc>'
            f'</w:tr>')
    # row1: vMerge restart with text in col0
    row1 = f'<w:tr>{cell0_restart()}{cell1("□行1の内容")}</w:tr>'
    # rows 2,3,4: vMerge continuation in col0
    row2 = f'<w:tr>{cell0_continue()}{cell1("□行2の内容")}</w:tr>'
    row3 = f'<w:tr>{cell0_continue()}{cell1("□行3の内容")}</w:tr>'
    row4 = f'<w:tr>{cell0_continue()}{cell1("□行4の内容")}</w:tr>'

    table = (f'<w:tbl>'
             f'<w:tblPr><w:tblW w:w="{tcw_col1 + tcw_col2}" w:type="dxa"/>{borders}</w:tblPr>'
             f'<w:tblGrid><w:gridCol w:w="{tcw_col1}"/><w:gridCol w:w="{tcw_col2}"/></w:tblGrid>'
             f'{row0}{row1}{row2}{row3}{row4}'
             f'</w:tbl>')
    doc = (f'<?xml version="1.0"?>'
           f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{table}{sect}</w:body></w:document>')
    return doc


def build(label, sz, charspace, tcw, text="組織的管理措置"):
    doc = make_doc(sz, charspace, tcw, text)
    path = os.path.join(OUT_DIR, f"{label}.docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc)
    print(f"Built {path}")


# CONFIRM baseline: matches b35123 (sz=21=10.5pt, cs=-2714, tcW=1271). Expect 4 chars / 12.45pt
build("B_baseline_match_b35", 21, -2714, 1271)

# Vary charSpace at fixed fs=21, tcW=1271
for cs in [0, -500, -1000, -2000, -2714, -3500, -4500]:
    sign = 'p' if cs >= 0 else 'n'
    build(f"P_fs21_cs{sign}{abs(cs):05d}_tcw1271", 21, cs, 1271)

# Vary fs at fixed cs=-2714, tcW=1271
for sz in [18, 24, 28]:  # 9pt, 12pt, 14pt
    build(f"P_fs{sz:02d}_csn02714_tcw1271", sz, -2714, 1271)

# Vary tcW at fixed fs=21, cs=-2714
for tcw in [800, 1000, 1500, 2000, 2500]:
    build(f"P_fs21_csn02714_tcw{tcw:04d}", 21, -2714, tcw)

# Combination: large vary
for sz in [18, 21, 24]:
    for cs in [0, -1500, -3000]:
        build(f"C_fs{sz:02d}_cs{('p' if cs>=0 else 'n')}{abs(cs):05d}", sz, cs, 1271)

print("Done.")
