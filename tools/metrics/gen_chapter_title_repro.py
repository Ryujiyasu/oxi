"""S188: build 4-variant minimal repro for the 14pt bold MS Mincho TOC
chapter title line height question, isolating the over-allocation
identified in S187 (3a4f drift origin).

Each variant: 5 body paragraphs (11pt MS Mincho regular) +
1 chapter-title paragraph (varies) + 5 more body paragraphs.

Variants:
  CT_A_14bold_MS    14pt bold MS Mincho             (the actual 3a4f pattern)
  CT_B_11_MS        11pt regular MS Mincho           (size control)
  CT_C_14_MS        14pt regular (not bold)          (bold control)
  CT_D_14bold_Arial 14pt bold Arial                  (font control)

Output: tools/golden-test/repros/chapter_title/CT_*.docx
Run:    python tools/metrics/gen_chapter_title_repro.py
"""
from __future__ import annotations
import os, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent.parent
OUT = REPO / 'tools' / 'golden-test' / 'repros' / 'chapter_title'
OUT.mkdir(parents=True, exist_ok=True)

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>'''
RELS_ROOT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''

NUMBERING = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="33">
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimalFullWidth"/>
      <w:lvlRestart w:val="0"/>
      <w:pStyle w:val="3"/>
      <w:suff w:val="space"/>
      <w:lvlText w:val="【第%3条"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="57" w:hanging="57"/></w:pPr>
      <w:rPr><w:rFonts w:hint="eastAsia"/><w:b/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="50"><w:abstractNumId w:val="33"/></w:num>
</w:numbering>'''
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
    </w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="a">
    <w:name w:val="Normal"/>
    <w:rPr><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="a"/>
    <w:next w:val="a"/>
    <w:qFormat/>
    <w:pPr>
      <w:numPr><w:ilvl w:val="2"/><w:numId w:val="50"/></w:numPr>
      <w:outlineLvl w:val="2"/>
    </w:pPr>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/></w:rPr>
  </w:style>
</w:styles>'''
SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="compressPunctuation"/>
</w:settings>'''


def body_para(i: int) -> str:
    return ('<w:p><w:pPr><w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
            '<w:sz w:val="22"/></w:rPr></w:pPr>'
            '<w:r><w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
            '<w:sz w:val="22"/></w:rPr>'
            f'<w:t>本文段落{i}この文章で line height を測定します。</w:t>'
            '</w:r></w:p>')


def chapter_title(sz: int, bold: bool, font: str) -> str:
    bold_tag = '<w:b/>' if bold else ''
    return (f'<w:p><w:pPr><w:rPr>'
            f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>'
            f'{bold_tag}'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
            f'</w:rPr></w:pPr>'
            f'<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="eastAsia"/>'
            f'{bold_tag}'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
            f'<w:t>第９章　無期労働契約への転換</w:t>'
            f'</w:r></w:p>')


def empty_heading3() -> str:
    """Mimics 3a4f wi=173: pStyle=3 with one fullwidth space."""
    return ('<w:p><w:pPr><w:pStyle w:val="3"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">　</w:t></w:r></w:p>')


def empty_inline_sz28_bold() -> str:
    """S194: matches 3a4f's actual pi=132 pattern — empty paragraph
    with inline pPr/rPr sz=28 bold + small ind. No pStyle."""
    return ('<w:p><w:pPr>'
            '<w:ind w:left="210" w:hanging="210"/>'
            '<w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
            '</w:pPr></w:p>')


def make_doc_body(sz: int, bold: bool, font: str, with_h3: bool = False, with_inline_sz28: bool = False) -> str:
    paras = []
    for i in range(1, 6):
        paras.append(body_para(i))
    if with_h3:
        paras.append(empty_heading3())
    if with_inline_sz28:
        paras.append(empty_inline_sz28_bold())
    paras.append(chapter_title(sz, bold, font))
    for i in range(6, 11):
        paras.append(body_para(i))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
    body = ''.join(paras) + sect
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{body}</w:body></w:document>')
    return doc


VARIANTS = [
    # (label, sz, bold, font, with_empty_heading3, with_inline_sz28_empty)
    ('A_14bold_MS',          28, True,  'ＭＳ 明朝', False, False),
    ('B_11_MS',              22, False, 'ＭＳ 明朝', False, False),
    ('C_14_MS',              28, False, 'ＭＳ 明朝', False, False),
    ('D_14bold_Arial',       28, True,  'Arial',    False, False),
    # S188 P-B extension: with empty heading-3 before chapter (mistaken hypothesis)
    ('E_14bold_MS_h3',       28, True,  'ＭＳ 明朝', True,  False),
    ('F_11_MS_h3',           22, False, 'ＭＳ 明朝', True,  False),
    # S189: same as E/F but heading-3 has full numPr
    ('G_14bold_MS_h3_numPr', 28, True,  'ＭＳ 明朝', True,  False),
    ('H_11_MS_h3_numPr',     22, False, 'ＭＳ 明朝', True,  False),
    # S194: CORRECT 3a4f pattern — empty paragraph with INLINE pPr/rPr sz=28 bold (no pStyle)
    ('I_inline_sz28_empty',  28, True,  'ＭＳ 明朝', False, True),
    ('J_inline_sz28_empty_11ch', 22, False, 'ＭＳ 明朝', False, True),
]


def write_docx(label: str, sz: int, bold: bool, font: str, with_h3: bool = False, with_inline_sz28: bool = False):
    doc_xml = make_doc_body(sz, bold, font, with_h3, with_inline_sz28)
    path = OUT / f'CT_{label}.docx'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', RELS_ROOT)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/numbering.xml', NUMBERING)
        z.writestr('word/document.xml', doc_xml.encode('utf-8'))
    print(f'  wrote {path}')


def main():
    print(f'Writing {len(VARIANTS)} variants to {OUT}')
    for v in VARIANTS:
        write_docx(*v)


if __name__ == '__main__':
    main()
