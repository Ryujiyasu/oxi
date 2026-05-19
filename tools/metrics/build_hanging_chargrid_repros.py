"""Hanging-indent + charGrid (linesAndChars docGrid) minimal repros.

Tests the hypothesis discovered in S109/S109c drilldown of 4a36b62:
- 4a36b62 has docGrid linesAndChars (charGrid active).
- Paragraph has w:ind left=368 hanging=160 (hanging indent, first_indent=-8pt @ sz=16).
- Oxi mod.rs:3678 sets effective_first_indent=0.0 for any first_indent when charGrid is active.
  Original d77a fix was for POSITIVE first_indent.
  For NEGATIVE first_indent (hanging), this loses the line-1 wrap extension.
- Oxi's line 1 wrap budget = available_width (463.4pt). Position is shifted to hanging
  (text starts -8pt early) but break_into_lines doesn't credit that 8pt back to line 1.
- Word appears to fit one MORE char on line 1, ending exactly at the right margin.

This repro creates a controlled set of hanging-indent paragraphs at varying font sizes
and hanging amounts, with text designed to overflow just at the right margin.

Variants:
  H1: sz=16 (8pt), hanging=160 (8pt = 1 char), CJK fill, ~58-60 chars/line
  H2: sz=21 (10.5pt), hanging=210 (10.5pt = 1 char), CJK fill
  H3: sz=24 (12pt), hanging=240 (12pt = 1 char), CJK fill
  H4: sz=16, hanging=320 (16pt = 2 chars), CJK fill
  H5: sz=16, NO hanging, plain indent for control (continuation-only baseline)
  H6: sz=16, hanging=160, NO charGrid (control: confirms no charGrid → wrap matches)
"""
import os
import zipfile

OUT_DIR = os.path.abspath("tools/metrics/hanging_chargrid_repro")
os.makedirs(OUT_DIR, exist_ok=True)

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

# 4a36b62 sectPr: pgSz 11904x16836, margins 1134/1134/1134/1134,
# docGrid linesAndChars linePitch=272 → ~38 chars/line at default font size.
SECT_CHARGRID = '<w:sectPr><w:pgSz w:w="11904" w:h="16836"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/><w:docGrid w:type="linesAndChars" w:linePitch="272"/></w:sectPr>'
SECT_NOGRID = '<w:sectPr><w:pgSz w:w="11904" w:h="16836"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'

# Wide enough CJK text to wrap 2-3 lines no matter the font size
LONG_CJK = "本報告書に記入された個人情報については、税務大学校との共同研究における国税庁保有行政記録情報利用における個票データ等の利用に関する業務のみに使用し、利用者の許可なくこれら以外の目的で使用しない。"
# 99 chars

def para(rpr, ind_attrs, text):
    return (f'<w:p><w:pPr>'
            f'<w:spacing w:line="240" w:lineRule="exact"/>'
            f'<w:ind {ind_attrs}/>'
            f'<w:rPr>{rpr}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def build(label, sect, ppr_list):
    """ppr_list: list of (rpr_xml, ind_xml, text)."""
    paras = "\n".join(para(r, i, t) for r, i, t in ppr_list)
    doc = (f'<?xml version="1.0"?>'
           f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{paras}{sect}</w:body></w:document>')
    path = os.path.join(OUT_DIR, f"{label}.docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc)
    print(f"Built {path}")


RPR_8 = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:sz w:val="16"/><w:szCs w:val="16"/>'
RPR_105 = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:sz w:val="21"/><w:szCs w:val="21"/>'
RPR_12 = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/>'

# H1: 4a36b62 exact replica — sz=16, hanging=160 left=368, charGrid
build("H1_sz16_hang160_grid", SECT_CHARGRID, [
    (RPR_8, 'w:leftChars="99" w:left="368" w:hangingChars="100" w:hanging="160"', f"２　{LONG_CJK}"),
])

# H2: sz=21 (10.5pt), hanging=210
build("H2_sz21_hang210_grid", SECT_CHARGRID, [
    (RPR_105, 'w:left="420" w:hanging="210"', f"２　{LONG_CJK}"),
])

# H3: sz=24 (12pt), hanging=240
build("H3_sz24_hang240_grid", SECT_CHARGRID, [
    (RPR_12, 'w:left="480" w:hanging="240"', f"２　{LONG_CJK}"),
])

# H4: sz=16, hanging=320 (2-char hanging)
build("H4_sz16_hang320_grid", SECT_CHARGRID, [
    (RPR_8, 'w:left="528" w:hanging="320"', f"２．{LONG_CJK}"),
])

# H5: sz=16, plain left indent (no hanging, no firstLine) — continuation-only baseline
build("H5_sz16_plain_grid", SECT_CHARGRID, [
    (RPR_8, 'w:left="368"', f"２　{LONG_CJK}"),
])

# H6: sz=16, hanging=160, NO charGrid — control: does Oxi/Word agree without charGrid?
build("H6_sz16_hang160_nogrid", SECT_NOGRID, [
    (RPR_8, 'w:left="368" w:hanging="160"', f"２　{LONG_CJK}"),
])

# H7: sz=16, POSITIVE first_indent=160 (regular indent), charGrid — d77a-style control
# This MUST keep current Oxi behavior (effective=0 for positive first_indent).
build("H7_sz16_firstLine160_grid", SECT_CHARGRID, [
    (RPR_8, 'w:left="368" w:firstLine="160"', f"２　{LONG_CJK}"),
])

print("Done.")
