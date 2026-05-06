# -*- coding: utf-8 -*-
"""
Per-char advance minimal repros to isolate Bug B residual contributors.

Variants (all written to tools/metrics/_repros/):
- V1: 10.5pt MS Mincho, cs=0, no kern              (baseline)
- V2: 10.5pt MS Mincho, cs=-9 (raw -0.45pt), no kern
- V3: 10.5pt MS Mincho, cs=0,  kern=2 (active)
- V4: 10.5pt MS Mincho, cs=-9, kern=2
- V5: V4 + autoSpaceDE=0 + autoSpaceDN=0
- V6: V5 + snapToGrid=0
- V7: V6 + jc=both (justify)

Each doc has the same single-line CJK text.
"""
import os, zipfile
from pathlib import Path

OUT = Path('tools/metrics/_repros')
OUT.mkdir(parents=True, exist_ok=True)

CT = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

CT_NO_SETTINGS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

DOC_RELS_WITH_SETTINGS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

def make_settings(*, balance_sb_db=False, use_fe_layout=False, char_spacing_control=None,
                  do_not_expand_shift_return=False, balance_only=False, fe_only=False,
                  csc_only=False):
    """Build a minimal settings.xml. Each toggle adds the corresponding flag."""
    flags = []
    if balance_sb_db: flags.append('<w:balanceSingleByteDoubleByteWidth/>')
    if do_not_expand_shift_return: flags.append('<w:doNotExpandShiftReturn/>')
    if use_fe_layout: flags.append('<w:useFELayout/>')
    csc = ''
    if char_spacing_control:
        csc = f'<w:characterSpacingControl w:val="{char_spacing_control}"/>'
    body = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
{csc}
<w:compat>
{''.join(flags)}
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""
    return body.encode('utf-8')

def styles_xml(default_kern=None):
    kern = f'<w:kern w:val="{default_kern}"/>' if default_kern is not None else ''
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>{kern}<w:sz w:val="21"/></w:rPr></w:rPrDefault>
</w:docDefaults>
</w:styles>""".encode('utf-8')

# Sample text: 10 CJK chars, all fullwidth. Avoid yakumono.
TEXT = '匿名データの名称、年次等'  # 12 chars, the 1636 item 1 text

def doc_xml(cs, snapToGrid=True, autoSpaceDE=True, autoSpaceDN=True, jc=None,
            no_grid_section=False):
    rpr_parts = ['<w:rFonts w:hint="eastAsia"/>']
    if cs is not None and cs != 0:
        rpr_parts.append(f'<w:spacing w:val="{cs}"/>')
    rpr_parts.append('<w:sz w:val="21"/>')
    rpr_xml = '<w:rPr>' + ''.join(rpr_parts) + '</w:rPr>'

    ppr_parts = []
    if not snapToGrid:
        ppr_parts.append('<w:snapToGrid w:val="0"/>')
    if not autoSpaceDE:
        ppr_parts.append('<w:autoSpaceDE w:val="0"/>')
    if not autoSpaceDN:
        ppr_parts.append('<w:autoSpaceDN w:val="0"/>')
    if jc is not None:
        ppr_parts.append(f'<w:jc w:val="{jc}"/>')
    ppr_parts.append('<w:spacing w:line="240" w:lineRule="exact"/>')
    ppr_xml = '<w:pPr>' + ''.join(ppr_parts) + '</w:pPr>' if ppr_parts else ''

    para = f'<w:p>{ppr_xml}<w:r>{rpr_xml}<w:t>{TEXT}</w:t></w:r></w:p>'

    if no_grid_section:
        sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
                'w:header="720" w:footer="720" w:gutter="0"/>'
                '<w:cols w:space="425"/></w:sectPr>')
    else:
        sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
                'w:header="720" w:footer="720" w:gutter="0"/>'
                '<w:cols w:space="425"/>'
                '<w:docGrid w:type="linesAndChars" w:linePitch="272"/>'
                '</w:sectPr>')

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{para}{sect}</w:body></w:document>').encode('utf-8')

def build(name, *, cs=0, kern=None, snapToGrid=True, autoSpaceDE=True, autoSpaceDN=True,
          jc=None, no_grid_section=False):
    out = OUT / name
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', styles_xml(default_kern=kern))
        z.writestr('word/document.xml', doc_xml(cs, snapToGrid, autoSpaceDE, autoSpaceDN, jc, no_grid_section))
    print(f'wrote {out}')

def doc_xml_table(cs, snapToGrid=True, autoSpaceDE=True, autoSpaceDN=True, jc=None,
                  style_cs=None, line_pitch=272):
    """Variant placing the test paragraph inside a single-cell table to mimic 1636."""
    style_cs_xml = f'<w:spacing w:val="{style_cs}"/>' if style_cs is not None else ''
    style_xml = f"""<w:style w:type="paragraph" w:customStyle="1" w:styleId="a3"><w:name w:val="一太郎"/><w:pPr><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:jc w:val="both"/></w:pPr><w:rPr>{style_cs_xml}<w:sz w:val="21"/></w:rPr></w:style>"""
    rpr_parts = ['<w:rFonts w:hint="eastAsia"/>']
    if cs is not None and cs != 0:
        rpr_parts.append(f'<w:spacing w:val="{cs}"/>')
    rpr_parts.append('<w:sz w:val="21"/>')
    rpr_xml = '<w:rPr>' + ''.join(rpr_parts) + '</w:rPr>'

    ppr_parts = [f'<w:pStyle w:val="a3"/>']
    if not snapToGrid:
        ppr_parts.append('<w:snapToGrid w:val="0"/>')
    if jc is not None:
        ppr_parts.append(f'<w:jc w:val="{jc}"/>')
    ppr_parts.append('<w:spacing w:line="240" w:lineRule="exact"/>')
    ppr_xml = '<w:pPr>' + ''.join(ppr_parts) + '</w:pPr>'

    para = f'<w:p>{ppr_xml}<w:r>{rpr_xml}<w:t>{TEXT}</w:t></w:r></w:p>'

    # Wrap in a single full-width cell mimicking 1636's gridSpan=3 cell
    table = f'''<w:tbl>
      <w:tblPr><w:tblW w:w="9923" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="9923"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:w="9923" w:type="dxa"/></w:tcPr>{para}</w:tc></w:tr>
    </w:tbl><w:p/>'''

    sect = (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            f'w:header="720" w:footer="720" w:gutter="0"/>'
            f'<w:cols w:space="425"/>'
            f'<w:docGrid w:type="linesAndChars" w:linePitch="{line_pitch}"/>'
            f'</w:sectPr>')

    body = f'<w:body>{table}{sect}</w:body>'
    doc_main = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'{body}</w:document>')
    full_styles = f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii=\"ＭＳ 明朝\" w:eastAsia=\"ＭＳ 明朝\" w:hAnsi=\"ＭＳ 明朝\"/><w:kern w:val=\"2\"/><w:sz w:val=\"21\"/></w:rPr></w:rPrDefault>
</w:docDefaults>
{style_xml}
</w:styles>"""
    return doc_main.encode('utf-8'), full_styles.encode('utf-8')

def build_table(name, **kwargs):
    out = OUT / name
    doc, styles = doc_xml_table(**kwargs)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/document.xml', doc)
    print(f'wrote {out}')

# Existing 7 variants
build('repro_pcw_V1.docx', cs=0)
build('repro_pcw_V2.docx', cs=-9)
build('repro_pcw_V3.docx', cs=0, kern=2)
build('repro_pcw_V4.docx', cs=-9, kern=2)
build('repro_pcw_V5.docx', cs=-9, kern=2, autoSpaceDE=False, autoSpaceDN=False)
build('repro_pcw_V6.docx', cs=-9, kern=2, autoSpaceDE=False, autoSpaceDN=False, snapToGrid=False)
build('repro_pcw_V7.docx', cs=-9, kern=2, autoSpaceDE=False, autoSpaceDN=False, snapToGrid=False, jc='both')

# Bonus: V8 = no docGrid section (replicate Day 11 v0/v3 baseline)
build('repro_pcw_V8.docx', cs=0, no_grid_section=True)

# Table-cell variants mimicking 1636 item 3 context
build_table('repro_pcw_V9.docx',  cs=-9, snapToGrid=False)                       # bare table cell + cs=-9 + snap=0
build_table('repro_pcw_V10.docx', cs=-9, snapToGrid=False, style_cs=-1)         # + style cs=-1
build_table('repro_pcw_V11.docx', cs=-9, snapToGrid=False, style_cs=-1, jc='both')  # + jc=both


# Finding 3 investigation variants — incrementally add real-1636 properties to V11
# (real items 1-5 measure Word=9.5pt; V11 measures 10.0pt; +0.5pt residual unexplained)

def doc_xml_table_v2(*, cs, snapToGrid=False, autoSpaceDE=True, autoSpaceDN=True,
                     jc='both', style_cs=-1, line_pitch=272,
                     # NEW toggles for real-1636 properties
                     ind_left_chars=None, ind_left=None, ind_right=None,
                     wordWrap_para=None,            # None = inherit; True = <w:wordWrap/>
                     style_widowControl_off=False,
                     style_wordWrap_off=False,
                     style_adjustRightInd_off=False,
                     style_rFonts_explicit=False,    # ascii/hAnsi/cs all Mincho
                     multi_run=False,                # split text into 4 runs (some w/o hint)
                     spacing_before_lines=None,      # vertical-only, but real has it
                     ):
    style_widowControl = '<w:widowControl w:val="0"/>' if style_widowControl_off else ''
    style_wordWrap = '<w:wordWrap w:val="0"/>' if style_wordWrap_off else ''
    style_adjRI = '<w:adjustRightInd w:val="0"/>' if style_adjustRightInd_off else ''
    style_rFonts = '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="ＭＳ 明朝"/>' if style_rFonts_explicit else ''
    style_cs_xml = f'<w:spacing w:val="{style_cs}"/>' if style_cs is not None else ''
    style_xml = (
        f'<w:style w:type="paragraph" w:customStyle="1" w:styleId="a3">'
        f'<w:name w:val="一太郎"/>'
        f'<w:pPr>{style_widowControl}{style_wordWrap}<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>{style_adjRI}<w:jc w:val="both"/></w:pPr>'
        f'<w:rPr>{style_rFonts}{style_cs_xml}<w:sz w:val="21"/></w:rPr>'
        f'</w:style>'
    )

    # Build run(s)
    rpr_parts = ['<w:rFonts w:hint="eastAsia"/>']
    if cs is not None and cs != 0:
        rpr_parts.append(f'<w:spacing w:val="{cs}"/>')
    rpr_parts.append('<w:sz w:val="21"/>')
    rpr_xml_with_hint = '<w:rPr>' + ''.join(rpr_parts) + '</w:rPr>'

    # Same but without hint
    rpr_parts_no_hint = []
    if cs is not None and cs != 0:
        rpr_parts_no_hint.append(f'<w:spacing w:val="{cs}"/>')
    rpr_parts_no_hint.append('<w:sz w:val="21"/>')
    rpr_xml_no_hint = '<w:rPr>' + ''.join(rpr_parts_no_hint) + '</w:rPr>'

    if multi_run:
        # Split text "匿名データの名称、年次等" into 4 runs, second w/o hint
        runs = (
            f'<w:r>{rpr_xml_with_hint}<w:t>匿名データの</w:t></w:r>'
            f'<w:r>{rpr_xml_no_hint}<w:t>名称、</w:t></w:r>'
            f'<w:r>{rpr_xml_with_hint}<w:t>年次</w:t></w:r>'
            f'<w:r>{rpr_xml_with_hint}<w:t>等</w:t></w:r>'
        )
    else:
        runs = f'<w:r>{rpr_xml_with_hint}<w:t>{TEXT}</w:t></w:r>'

    ppr_parts = [f'<w:pStyle w:val="a3"/>']
    if wordWrap_para is True:
        ppr_parts.append('<w:wordWrap/>')
    elif wordWrap_para is False:
        ppr_parts.append('<w:wordWrap w:val="0"/>')
    if not snapToGrid:
        ppr_parts.append('<w:snapToGrid w:val="0"/>')
    spacing_attrs = []
    if spacing_before_lines is not None:
        spacing_attrs.append(f'w:beforeLines="{spacing_before_lines}" w:before="136"')
    spacing_attrs.append('w:line="240" w:lineRule="exact"')
    ppr_parts.append(f'<w:spacing {" ".join(spacing_attrs)}/>')
    ind_attrs = []
    if ind_left_chars is not None: ind_attrs.append(f'w:leftChars="{ind_left_chars}"')
    if ind_left is not None:       ind_attrs.append(f'w:left="{ind_left}"')
    if ind_right is not None:      ind_attrs.append(f'w:right="{ind_right}"')
    if ind_attrs:
        ppr_parts.append(f'<w:ind {" ".join(ind_attrs)}/>')
    if jc is not None:
        ppr_parts.append(f'<w:jc w:val="{jc}"/>')

    ppr_xml = '<w:pPr>' + ''.join(ppr_parts) + '</w:pPr>'
    para = f'<w:p>{ppr_xml}{runs}</w:p>'

    table = f'''<w:tbl>
      <w:tblPr><w:tblW w:w="9923" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="9923"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:w="9923" w:type="dxa"/></w:tcPr>{para}</w:tc></w:tr>
    </w:tbl><w:p/>'''

    sect = (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            f'w:header="720" w:footer="720" w:gutter="0"/>'
            f'<w:cols w:space="425"/>'
            f'<w:docGrid w:type="linesAndChars" w:linePitch="{line_pitch}"/>'
            f'</w:sectPr>')

    body = f'<w:body>{table}{sect}</w:body>'
    doc_main = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'{body}</w:document>')
    full_styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                   '<w:docDefaults><w:rPrDefault><w:rPr>'
                   '<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
                   '<w:kern w:val="2"/><w:sz w:val="21"/>'
                   '</w:rPr></w:rPrDefault></w:docDefaults>'
                   f'{style_xml}'
                   '</w:styles>')
    return doc_main.encode('utf-8'), full_styles.encode('utf-8')

def build_table_v2(name, **kwargs):
    out = OUT / name
    doc, styles = doc_xml_table_v2(**kwargs)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/document.xml', doc)
    print(f'wrote {out}')

# V12-V18: Each adds ONE real-1636 property to V11-equivalent baseline
build_table_v2('repro_pcw_V12.docx', cs=-9)                                              # V11 + 1636-style docDefault rFonts (Century-ish)
build_table_v2('repro_pcw_V13.docx', cs=-9, ind_left_chars=150, ind_left=315, ind_right=199)  # + indent
build_table_v2('repro_pcw_V14.docx', cs=-9, wordWrap_para=True)                          # + para-level wordWrap=true
build_table_v2('repro_pcw_V15.docx', cs=-9, style_widowControl_off=True, style_wordWrap_off=True, style_adjustRightInd_off=True)  # + missing style props
build_table_v2('repro_pcw_V16.docx', cs=-9, style_rFonts_explicit=True)                   # + explicit ascii/hAnsi/cs Mincho
build_table_v2('repro_pcw_V17.docx', cs=-9, multi_run=True)                              # + multi-run split (some w/o hint=eastAsia)
build_table_v2('repro_pcw_V18.docx', cs=-9,                                              # ALL real-1636 props combined
               ind_left_chars=150, ind_left=315, ind_right=199,
               wordWrap_para=True,
               style_widowControl_off=True, style_wordWrap_off=True, style_adjustRightInd_off=True,
               style_rFonts_explicit=True,
               multi_run=True,
               spacing_before_lines=50)

# V19+: Test settings.xml flags (the leading hypothesis for Finding 3 residual)
def build_table_with_settings(name, settings_kwargs, **doc_kwargs):
    out = OUT / name
    doc, styles = doc_xml_table_v2(**doc_kwargs)
    settings_xml = make_settings(**settings_kwargs)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS_WITH_SETTINGS)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/document.xml', doc)
        z.writestr('word/settings.xml', settings_xml)
    print(f'wrote {out}')

# V19: V11 + balanceSingleByteDoubleByteWidth alone
build_table_with_settings('repro_pcw_V19.docx', dict(balance_sb_db=True), cs=-9)
# V20: V11 + useFELayout alone
build_table_with_settings('repro_pcw_V20.docx', dict(use_fe_layout=True), cs=-9)
# V21: V11 + characterSpacingControl=compressPunctuation alone
build_table_with_settings('repro_pcw_V21.docx', dict(char_spacing_control='compressPunctuation'), cs=-9)
# V22: V11 + ALL three together
build_table_with_settings('repro_pcw_V22.docx',
                          dict(balance_sb_db=True, use_fe_layout=True,
                               char_spacing_control='compressPunctuation'),
                          cs=-9)
# V23: V18 (all para+style props) + ALL settings flags
build_table_with_settings('repro_pcw_V23.docx',
                          dict(balance_sb_db=True, use_fe_layout=True,
                               char_spacing_control='compressPunctuation',
                               do_not_expand_shift_return=True),
                          cs=-9,
                          ind_left_chars=150, ind_left=315, ind_right=199,
                          wordWrap_para=True,
                          style_widowControl_off=True, style_wordWrap_off=True, style_adjustRightInd_off=True,
                          style_rFonts_explicit=True,
                          multi_run=True,
                          spacing_before_lines=50)

# V24-V27: Isolate balance effect on base advance + cs interactions
build_table_with_settings('repro_pcw_V24.docx', dict(balance_sb_db=True), cs=0)      # balance + cs=0 (base advance only)
build_table_with_settings('repro_pcw_V25.docx', dict(balance_sb_db=True), cs=-20)    # balance + cs=-1pt
build_table_with_settings('repro_pcw_V26.docx', dict(balance_sb_db=True), cs=20)     # balance + cs=+1pt
build_table_with_settings('repro_pcw_V27.docx', dict(balance_sb_db=True), cs=-5)     # balance + cs=-0.25pt

print('Done.')
