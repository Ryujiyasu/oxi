"""Day 33 part 29 (2026-05-11) — Test widowControl=0 first-line overflow.

Hypothesis (Day 33 part 28): widowControl=0 paragraphs at page boundary
allow first-line "descender intrusion" into bottom margin in Word,
unlike widowControl=ON paragraphs (which strictly break per Day 33 part 1).

Variants:
- WO_widow_on:  fill paragraphs + multi-line test paragraph; widowControl=ON
- WO_widow_off: same but widowControl=OFF
Place test paragraph such that its first line would land at exactly the
page bottom (within 0.5pt). Measure Word's actual y position for the
test paragraph.
"""
from __future__ import annotations
import os, zipfile
from pathlib import Path

OUT = Path('tools/golden-test/repros/widow_overflow')
OUT.mkdir(parents=True, exist_ok=True)

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''


def make_styles(widow_control_value):
    """widow_control_value: '1' (ON) or '0' (OFF) for Normal style."""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/>
<w:sz w:val="21"/>
</w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="{widow_control_value}"/></w:pPr>
</w:style>
</w:styles>'''


def make_para(text='テスト'):
    return f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'


def make_multi_line_para(text):
    """A paragraph with text that wraps to multiple lines."""
    return f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'


def make_document(n_fill, test_text):
    # A4: 11906x16838 tw. 1in margins (1440tw=72pt each). Body content area
    # = 16838 - 1440*2 - footer = 13958 tw = 697.9 pt
    # docGrid linePitch=360tw=18pt → 38 lines per page (697.9 / 18)
    # We place n_fill single-line paragraphs to fill close to bottom,
    # then the test multi-line paragraph.
    fills = ''.join(make_para(f'F{i:02d}') for i in range(n_fill))
    test_para = make_multi_line_para(test_text)
    sect = ('<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:docGrid w:type="lines" w:linePitch="360"/>'
            '</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{fills}{test_para}{sect}</w:body></w:document>'


def build_docx(name, widow, n_fill, test_text):
    p = OUT / f'{name}.docx'
    if p.exists(): p.unlink()
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', make_styles(widow))
        z.writestr('word/document.xml', make_document(n_fill, test_text))
    print(f'  wrote {p}')


# Long Japanese text that wraps to ~3 lines at body width
LONG_TEXT = (
    '実際の文書では、ある段落がページの最後に到達し、次のページに継続することがあります。'
    'Word の動作はウィドウ・オーファン制御の有無によって変わる可能性があります。'
    'この最小再現テストでは、その境界における振る舞いを正確に測定します。'
)


if __name__ == '__main__':
    # 35 fill paragraphs at 18pt = 630pt content. + test para starts at y=72+630=702
    # 38 lines per page at linePitch=360 means line 38 starts at y=72+37*18=738.
    # So 35 fills place test_para start at y=72+35*18=702 → has room for 1 more line
    # at y=720 fitting (720+18=738 < 770).
    # 36 fills: test_para at y=72+36*18=720 → line 0 at 720 ends 738, fits.
    # 37 fills: test_para at y=72+37*18=738 → line 0 ends 756, fits.
    # 38 fills: test_para at y=72+38*18=756 → line 0 ends 774 > 770 OVERFLOWS.

    for widow_label, widow_val in [('on', '1'), ('off', '0')]:
        for n_fill in [36, 37, 38, 39]:
            name = f'WO_widow_{widow_label}_fill{n_fill}'
            build_docx(name, widow_val, n_fill, LONG_TEXT)
    print(f'Wrote 8 repros to {OUT}')
