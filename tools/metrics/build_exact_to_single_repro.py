"""S107 minimal repro: paragraph with lineRule=exact followed by Single paragraphs.

Tests d77a p.1 pi=1 (title, exact line=420) → pi=2 (empty Single) transition.
Word: gap 22.5pt
Oxi: gap 21pt
Diff: 1.5pt — Word's first Single line after exact has half-leading offset.

Variations:
  V1: pi=1 exact 420tw (21pt), pi=2 Single MS Gothic 12pt
  V2: V1 + pi=3 Single MS Gothic 12pt (to confirm constant after first transition)
  V3: pi=1 Single MS Gothic 14pt (no exact), pi=2 Single MS Gothic 12pt
  V4: pi=1 exact 420tw MS Gothic 14pt, pi=2 Single MS Mincho 10.5pt
"""
import os, zipfile
from pathlib import Path

OUT = Path('c:/Users/ryuji/oxi-main/tools/metrics/exact_to_single_repro')
OUT.mkdir(parents=True, exist_ok=True)

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:kern w:val="2"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a">
<w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
</w:style>
</w:styles>'''


def para_xml(text, font, font_sz, line_rule=None, line_val=None, ind=None):
    spacing = ''
    if line_rule:
        spacing = f'<w:spacing w:line="{line_val}" w:lineRule="{line_rule}"/>'
    ind_xml = f'<w:ind {ind}/>' if ind else ''
    runs = ''
    if text:
        runs = f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:hint="eastAsia"/><w:kern w:val="2"/><w:sz w:val="{font_sz}"/></w:rPr><w:t>{text}</w:t></w:r>'
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{font_sz}"/></w:rPr>'
    return f'<w:p><w:pPr>{spacing}{ind_xml}{rpr}</w:pPr>{runs}</w:p>'


def make_docx(name, body_paras):
    body = '\n'.join(body_paras) + '''
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>'''
    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>{body}</w:body></w:document>'''
    out_path = OUT / f"{name}.docx"
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc_xml)
    return out_path


def main():
    G = 'ＭＳ ゴシック'
    M = 'ＭＳ 明朝'
    # V1: pi=1 exact 420 (21pt) MS Gothic 12pt; pi=2 empty Single MS Gothic 12pt; pi=3 single line MS Gothic 12pt
    make_docx("V1_exact420_to_single", [
        para_xml('Exact21pt', G, 28, line_rule='exact', line_val='420'),
        para_xml('', G, 24),
        para_xml('LineAfter', G, 24),
    ])
    # V2: pi=1 exact 360 (18pt) MS Gothic 12pt
    make_docx("V2_exact360_to_single", [
        para_xml('Exact18pt', G, 28, line_rule='exact', line_val='360'),
        para_xml('', G, 24),
        para_xml('LineAfter', G, 24),
    ])
    # V3: no exact — both Single MS Gothic 12pt
    make_docx("V3_single_to_single", [
        para_xml('Single1', G, 24),
        para_xml('', G, 24),
        para_xml('Single3', G, 24),
    ])
    # V4: pi=1 exact 420 MS Gothic 14pt (matches d77a Heading 1); pi=2 empty MS Gothic 12pt; pi=3 line MS Gothic 12pt
    make_docx("V4_d77a_like_heading", [
        para_xml('Heading14pt', G, 28, line_rule='exact', line_val='420'),
        para_xml('', G, 24),
        para_xml('Body12pt', G, 24),
    ])
    # V5: exact at start, then multi-line para
    make_docx("V5_exact_then_multiline", [
        para_xml('H', G, 28, line_rule='exact', line_val='420'),
        para_xml('一二三四五六七八九十一二三四五六七八九十一二三四五六七八九十一二三四五六七八', G, 24),
    ])
    print(f"Built repros in {OUT}")


if __name__ == '__main__':
    main()
