"""Session 112 phase 3 — having isolated jc=both as the trigger, test
boundary conditions:

  v15_no_balanceSBDB    : jc=both + remove balanceSBDB
  v16_doNotCompress     : jc=both + characterSpacingControl=doNotCompress
  v17_no_csctl          : jc=both + no characterSpacingControl at all
  v18_no_compat15       : jc=both + no compatSetting=15
  v19_jc_no_neg9        : jc=both + no <w:spacing w:val="-9"/>
  v20_jc_kern_default   : jc=both + no <w:kern w:val="0"/> (default kerning)
  v21_jc_szCs_22        : jc=both + szCs=22 instead of 21
  v22_no_wordWrap       : jc=both + no <w:wordWrap w:val="0"/>
  v23_no_autoSpaceDE    : jc=both + no autoSpaceDE/DN
  v24_comma             : jc=both + content '，' (fullwidth comma)
  v25_kuten             : jc=both + content '。' (fullwidth full stop, kuten)
  v26_question          : jc=both + content '？'
  v27_jc_distribute     : jc=distribute (vs both)
  v28_jc_left           : jc=left (control — should NOT trigger)
"""
import os
import sys
import io
import zipfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.normpath(os.path.join(REPO, "tools/metrics/15076df_buildup/variants"))
os.makedirs(OUT_DIR, exist_ok=True)


def build_settings(balance_sbdb=True, char_space_ctl="compressPunctuation", compat_15=True):
    s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    s += '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
    s += '<w:compat>\n'
    if balance_sbdb:
        s += '<w:balanceSingleByteDoubleByteWidth/>\n'
    if compat_15:
        s += '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>\n'
    s += '</w:compat>\n'
    if char_space_ctl:
        s += f'<w:characterSpacingControl w:val="{char_space_ctl}"/>\n'
    s += '</w:settings>\n'
    return s.encode('utf-8')


def build_styles_with_jc(jc="both"):
    s = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:cs="ＭＳ 明朝"/>'
        '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/>'
        '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
        '</w:rPr></w:rPrDefault></w:docDefaults>\n'
    )
    if jc:
        s += (f'<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              f'<w:name w:val="Normal"/><w:qFormat/>'
              f'<w:pPr><w:jc w:val="{jc}"/></w:pPr></w:style>\n')
    else:
        s += '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style>\n'
    s += '</w:styles>\n'
    return s.encode('utf-8')


def build_document(
    text="１．提供を受けた匿名データの名称",
    cs_val=-9,
    kern_val=0,
    sz_cs=21,
    word_wrap=True,
    auto_space=True,
):
    cs_xml = f'<w:spacing w:val="{cs_val}"/>' if cs_val is not None else ""
    kern_xml = f'<w:kern w:val="{kern_val}"/>' if kern_val is not None else ""
    szcs_xml = f'<w:szCs w:val="{sz_cs}"/>' if sz_cs is not None else ""
    ww_xml = '<w:wordWrap w:val="0"/>' if word_wrap else ""
    asp_xml = '<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>' if auto_space else ""
    s = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        '<w:body>\n'
        '<w:tbl><w:tblPr><w:tblW w:w="1968" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="12" w:type="dxa"/><w:right w:w="12" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="1968"/></w:tblGrid>'
        '<w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="1968" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr>'
        f'{ww_xml}{asp_xml}'
        '<w:adjustRightInd w:val="0"/>'
        '<w:spacing w:line="240" w:lineRule="exact"/>'
        '<w:ind w:left="215" w:right="76" w:hanging="192"/>'
        '</w:pPr>'
        '<w:r><w:rPr>'
        '<w:rFonts w:hAnsi="ＭＳ 明朝" w:cs="ＭＳ 明朝" w:hint="eastAsia"/>'
        f'{cs_xml}{kern_xml}{szcs_xml}'
        '</w:rPr>'
        f'<w:t>{text}</w:t></w:r>'
        '</w:p></w:tc></w:tr></w:tbl>\n'
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="851" w:right="1134" w:bottom="567" w:left="1134"'
        ' w:header="851" w:footer="567" w:gutter="0"/>'
        '<w:docGrid w:type="lines" w:linePitch="336"/>'
        '</w:sectPr>\n'
        '</w:body></w:document>\n'
    )
    return s.encode('utf-8')


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
    '<Default Extension="xml" ContentType="application/xml"/>\n'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
    '</Types>\n'
).encode('utf-8')

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n'
    '</Relationships>\n'
).encode('utf-8')

DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n'
    '</Relationships>\n'
).encode('utf-8')


def write_docx(name, styles, document, settings):
    out_path = os.path.join(OUT_DIR, f"{name}.docx")
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/settings.xml', settings)
    print(f"wrote {out_path}")


def main():
    jc_styles = build_styles_with_jc("both")
    default_settings = build_settings()
    default_doc = build_document()

    # v15: jc=both, no balanceSBDB
    write_docx("v15_no_balanceSBDB", jc_styles, default_doc,
               build_settings(balance_sbdb=False))
    # v16: jc=both, doNotCompress
    write_docx("v16_doNotCompress", jc_styles, default_doc,
               build_settings(char_space_ctl="doNotCompress"))
    # v17: jc=both, no characterSpacingControl
    write_docx("v17_no_csctl", jc_styles, default_doc,
               build_settings(char_space_ctl=None))
    # v18: jc=both, no compatibilityMode=15
    write_docx("v18_no_compat15", jc_styles, default_doc,
               build_settings(compat_15=False))
    # v19: jc=both, no cs=-9
    write_docx("v19_jc_no_neg9", jc_styles,
               build_document(cs_val=None), default_settings)
    # v20: jc=both, no kern=0 (default kerning)
    write_docx("v20_jc_kern_default", jc_styles,
               build_document(kern_val=None), default_settings)
    # v21: jc=both, szCs=22
    write_docx("v21_jc_szCs_22", jc_styles,
               build_document(sz_cs=22), default_settings)
    # v22: jc=both, no wordWrap=0
    write_docx("v22_no_wordWrap", jc_styles,
               build_document(word_wrap=False), default_settings)
    # v23: jc=both, no autoSpace
    write_docx("v23_no_autoSpaceDE", jc_styles,
               build_document(auto_space=False), default_settings)
    # v24: jc=both + content with comma
    write_docx("v24_comma", jc_styles,
               build_document(text="１，提供を受けた匿名データの名称"), default_settings)
    # v25: jc=both + content with kuten
    write_docx("v25_kuten", jc_styles,
               build_document(text="１。提供を受けた匿名データの名称"), default_settings)
    # v26: jc=both + content with question mark
    write_docx("v26_question", jc_styles,
               build_document(text="１？提供を受けた匿名データの名称"), default_settings)
    # v27: jc=distribute
    write_docx("v27_jc_distribute", build_styles_with_jc("distribute"),
               default_doc, default_settings)
    # v28: jc=left (control — should NOT trigger)
    write_docx("v28_jc_left", build_styles_with_jc("left"),
               default_doc, default_settings)


if __name__ == "__main__":
    main()
