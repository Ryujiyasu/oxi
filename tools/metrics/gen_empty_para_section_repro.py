"""S150: Build TR_V400 minimal repros for a1d6's empty-para-before-section-header
+11.55pt drift pattern.

a1d6 OOXML pattern for section transitions:
  <w:p ...empty.../>             (no spacing, ls=12 lsr=0 default single)
  <w:p ...section header...>     (sb=146-154, ls=14exact or 12.95exact, sz=18-20)

Drift: Oxi places section header +11.55pt later than Word.

Variants:
  V400a: empty para + section header (sb=146 line=280 exact, sz=20)
  V400b: same but no preceding empty para (sanity)
  V400c: 2 empty paras + section header
  V400d: empty para WITH explicit line=240 exact + section header
  V400e: empty para has same pStyle as section header (inherit)
  V400f: empty para inside table cell + section header outside
  V400g: section header has beforeAutoSpacing
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'empty_para_section')

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

# Match a1d6's pStyle "ac" + settings.xml
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="ac">
<w:name w:val="ac"/>
<w:pPr><w:widowControl w:val="0"/><w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/><w:spacing w:line="210" w:lineRule="exact"/><w:jc w:val="both"/></w:pPr>
<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="ＭＳ 明朝"/><w:spacing w:val="-1"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
</w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:useFELayout/>
<w:adjustLineHeightInTable/>
<w:doNotExpandShiftReturn/>
</w:compat>
</w:settings>'''


def empty_para(line: int = 240, line_rule: str | None = 'exact', sz: int = 20) -> str:
    spacing = f'<w:spacing w:line="{line}" w:lineRule="{line_rule}"/>' if line_rule else f'<w:spacing w:line="{line}"/>'
    return f'<w:p><w:pPr>{spacing}</w:pPr></w:p>'


def empty_para_default() -> str:
    """No pPr - default single spacing."""
    return '<w:p/>'


def section_header(text: str = '５　匿名データの提供を受ける方法及び提供希望年月日',
                   sb: int = 146, sb_lines: int | None = 50,
                   line: int = 280, sz: int = 20) -> str:
    sb_attrs = f'w:before="{sb}"'
    if sb_lines is not None:
        sb_attrs = f'w:beforeLines="{sb_lines}" w:before="{sb}"'
    return ('<w:p><w:pPr>'
            '<w:pStyle w:val="ac"/>'
            f'<w:spacing {sb_attrs} w:line="{line}" w:lineRule="exact"/>'
            f'<w:rPr><w:sz w:val="{sz}"/></w:rPr>'
            '</w:pPr>'
            '<w:r>'
            f'<w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="{sz}"/></w:rPr>'
            f'<w:t>{text}</w:t>'
            '</w:r>'
            '</w:p>')


def body_marker(text: str = 'MARKER') -> str:
    return f'<w:p><w:pPr><w:pStyle w:val="ac"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


def doc_xml(body: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body}
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1247" w:right="1077" w:bottom="1440" w:left="1077" w:header="851" w:footer="992" w:gutter="0"/>
<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/>
</w:sectPr>
</w:body>
</w:document>"""


def write_docx(label: str, doc: str):
    out = os.path.join(OUT_DIR, f'{label}.docx')
    os.makedirs(OUT_DIR, exist_ok=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', DOC_RELS)
        zf.writestr('word/settings.xml', SETTINGS)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/document.xml', doc)
    return out


def main():
    # V400a: BEFORE marker + empty para + section header (the a1d6 pattern)
    body_a = (body_marker('BEFORE') + empty_para_default() +
              section_header() + body_marker('AFTER'))
    write_docx('V400a_empty_default_then_section', doc_xml(body_a))

    # V400b: no empty para (control)
    body_b = body_marker('BEFORE') + section_header() + body_marker('AFTER')
    write_docx('V400b_no_empty_para', doc_xml(body_b))

    # V400c: 2 empty paras
    body_c = (body_marker('BEFORE') + empty_para_default() + empty_para_default() +
              section_header() + body_marker('AFTER'))
    write_docx('V400c_two_empty_paras', doc_xml(body_c))

    # V400d: empty para with explicit line=240 exact
    body_d = (body_marker('BEFORE') + empty_para(line=240, line_rule='exact') +
              section_header() + body_marker('AFTER'))
    write_docx('V400d_empty_exact240', doc_xml(body_d))

    # V400e: empty para inherits pStyle "ac"
    body_e = (body_marker('BEFORE') +
              '<w:p><w:pPr><w:pStyle w:val="ac"/></w:pPr></w:p>' +
              section_header() + body_marker('AFTER'))
    write_docx('V400e_empty_pstyle_ac', doc_xml(body_e))

    # V400f: section header without beforeLines (just sb=146)
    body_f = (body_marker('BEFORE') + empty_para_default() +
              section_header(sb_lines=None) + body_marker('AFTER'))
    write_docx('V400f_no_beforeLines', doc_xml(body_f))

    # V400g: empty para line=0 (= "auto")
    body_g = (body_marker('BEFORE') + empty_para(line=0, line_rule='auto') +
              section_header() + body_marker('AFTER'))
    write_docx('V400g_empty_auto0', doc_xml(body_g))

    print('Done. Repros in', OUT_DIR)


if __name__ == '__main__':
    main()
