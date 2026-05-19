"""V2: Hanging + charGrid + compressPunctuation + compat=15 (4a36b62-style settings).

H1 v1 vs 4a36b62 actual differed by 2 chars Word-side. Hypothesis: compressPunctuation
contributes the +2. This v2 reproduces 4a36b62's settings exactly to isolate the effect.
"""
import os, zipfile

OUT_DIR = os.path.abspath("tools/metrics/hanging_chargrid_repro_v2")
os.makedirs(OUT_DIR, exist_ok=True)

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>'''

SECT_GRID = '<w:sectPr><w:pgSz w:w="11904" w:h="16836"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/><w:docGrid w:type="linesAndChars" w:linePitch="272"/></w:sectPr>'

LONG_CJK = "本報告書に記入された個人情報については、税務大学校との共同研究における国税庁保有行政記録情報利用における個票データ等の利用に関する業務のみに使用し、利用者の許可なくこれら以外の目的で使用しない。"

RPR_8 = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:sz w:val="16"/><w:szCs w:val="16"/>'
RPR_105 = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:sz w:val="21"/><w:szCs w:val="21"/>'
RPR_12 = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/>'


def para(rpr, ind_attrs, text):
    return (f'<w:p><w:pPr>'
            f'<w:spacing w:line="240" w:lineRule="exact"/>'
            f'<w:ind {ind_attrs}/>'
            f'<w:rPr>{rpr}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def build(label, ppr_list):
    paras = "\n".join(para(r, i, t) for r, i, t in ppr_list)
    doc = (f'<?xml version="1.0"?>'
           f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{paras}{SECT_GRID}</w:body></w:document>')
    path = os.path.join(OUT_DIR, f"{label}.docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc)
    print(f"Built {path}")


# V2 repros — same as v1 but with compressPunctuation + compat15
build("H1v2_sz16_hang160_grid", [
    (RPR_8, 'w:leftChars="99" w:left="368" w:hangingChars="100" w:hanging="160"', f"２　{LONG_CJK}"),
])
build("H2v2_sz21_hang210_grid", [
    (RPR_105, 'w:left="420" w:hanging="210"', f"２　{LONG_CJK}"),
])
build("H3v2_sz24_hang240_grid", [
    (RPR_12, 'w:left="480" w:hanging="240"', f"２　{LONG_CJK}"),
])
build("H4v2_sz16_hang320_grid", [
    (RPR_8, 'w:left="528" w:hanging="320"', f"２．{LONG_CJK}"),
])
build("H5v2_sz16_plain_grid", [
    (RPR_8, 'w:left="368"', f"２　{LONG_CJK}"),
])
build("H7v2_sz16_firstLine160_grid", [
    (RPR_8, 'w:left="368" w:firstLine="160"', f"２　{LONG_CJK}"),
])

# Extra: vary hanging amount to detect cap (does Word allow hanging extension > 1 char?)
build("H8v2_sz16_hang80_grid", [   # half-char hanging
    (RPR_8, 'w:left="368" w:hanging="80"', f"２　{LONG_CJK}"),
])
build("H9v2_sz16_hang240_grid", [  # 1.5-char hanging
    (RPR_8, 'w:left="368" w:hanging="240"', f"２　{LONG_CJK}"),
])
build("H10v2_sz16_hang480_grid", [  # 3-char hanging
    (RPR_8, 'w:left="528" w:hanging="480"', f"２　{LONG_CJK}"),
])

print("Done.")
