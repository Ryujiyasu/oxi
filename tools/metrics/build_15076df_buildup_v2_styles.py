"""Session 112 phase 2 — bisect WITHIN 15076df styles.xml to find the
specific element that causes '．' = 6.0pt.

v3_styles (full styles.xml swapped in) reproduces the bug. The styles.xml
differs from minimal in:
  (a) docDefaults uses theme fonts (asciiTheme=minorHAnsi etc.) vs literal
      ＭＳ 明朝
  (b) defines named styles 'a' (Normal) with widowControl=0, jc=both,
      ＭＳ 明朝 ascii+eastAsia
  (c) defines Table Grid 'a3' with ind ChrBased + jc=both + borders
  (d) defines 14 more styles (header/footer/balloon/comment — unlikely)

Test each suspect IN ISOLATION on top of minimal base:
  v8_jc_both       : Normal-style 'a' adds jc=both only
  v9_widow         : Normal-style 'a' adds widowControl=0 only
  v10_normal_full  : Normal-style 'a' full (widow+jc+font)
  v11_themefonts   : docDefaults uses theme fonts (no literal)
  v12_tableGrid    : Table Grid 'a3' added + document references tblStyle=a3
  v13_jc_doc       : minimal styles, document.xml adds <w:jc w:val="both"/>
  v14_widow_doc    : minimal styles, document.xml adds <w:widowControl w:val="0"/>
"""
import os
import sys
import io
import zipfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.normpath(os.path.join(REPO, "tools/metrics/15076df_buildup/variants"))
os.makedirs(OUT_DIR, exist_ok=True)


MIN_SETTINGS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
    b'<w:compat>\n'
    b'<w:balanceSingleByteDoubleByteWidth/>\n'
    b'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>\n'
    b'</w:compat>\n'
    b'<w:characterSpacingControl w:val="compressPunctuation"/>\n'
    b'</w:settings>\n'
)


def build_styles(extra_styles: str = "", theme_fonts: bool = False) -> bytes:
    if theme_fonts:
        rfonts = ('<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"'
                  ' w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>')
    else:
        rfonts = ('<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
                  ' w:eastAsia="ＭＳ 明朝" w:cs="ＭＳ 明朝"/>')
    s = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        f'{rfonts}'
        '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/>'
        '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
        '</w:rPr></w:rPrDefault></w:docDefaults>\n'
    )
    if not extra_styles:
        # Default Normal placeholder
        s += '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style>\n'
    else:
        s += extra_styles
    s += '</w:styles>\n'
    return s.encode('utf-8')


def build_document(extra_ppr: str = "", tbl_style: str = "") -> bytes:
    tbl_style_xml = f'<w:tblStyle w:val="{tbl_style}"/>' if tbl_style else ""
    s = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        '<w:body>\n'
        f'<w:tbl><w:tblPr>{tbl_style_xml}<w:tblW w:w="1968" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="12" w:type="dxa"/><w:right w:w="12" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="1968"/></w:tblGrid>'
        '<w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="1968" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr>'
        '<w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>'
        '<w:adjustRightInd w:val="0"/>'
        '<w:spacing w:line="240" w:lineRule="exact"/>'
        '<w:ind w:left="215" w:right="76" w:hanging="192"/>'
        f'{extra_ppr}'
        '</w:pPr>'
        '<w:r><w:rPr>'
        '<w:rFonts w:hAnsi="ＭＳ 明朝" w:cs="ＭＳ 明朝" w:hint="eastAsia"/>'
        '<w:spacing w:val="-9"/><w:kern w:val="0"/><w:szCs w:val="21"/>'
        '</w:rPr>'
        '<w:t>１．提供を受けた匿名データの名称</w:t></w:r>'
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


def write_docx(name, styles, document):
    out_path = os.path.join(OUT_DIR, f"{name}.docx")
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/settings.xml', MIN_SETTINGS)
    print(f"wrote {out_path}")


def main():
    # v8: Normal style 'a' with ONLY jc=both
    s = '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>\n'
    write_docx("v8_jc_both", build_styles(s), build_document())

    # v9: Normal style 'a' with ONLY widowControl=0
    s = '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>\n'
    write_docx("v9_widow", build_styles(s), build_document())

    # v10: Normal style 'a' full (widow + jc + font) — matches 15076df Normal exactly
    s = ('<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
         '<w:qFormat/><w:rsid w:val="00010BC0"/>'
         '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>'
         '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/></w:rPr></w:style>\n')
    write_docx("v10_normal_full", build_styles(s), build_document())

    # v11: docDefaults uses theme fonts (no literal eastAsia=ＭＳ 明朝)
    write_docx("v11_themefonts", build_styles(theme_fonts=True), build_document())

    # v12: Add Table Grid 'a3' + document references tblStyle=a3
    s = ('<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style>\n'
         '<w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/>'
         '<w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
         '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>\n'
         '<w:style w:type="table" w:styleId="a3"><w:name w:val="Table Grid"/><w:basedOn w:val="a1"/>'
         '<w:pPr><w:spacing w:line="254" w:lineRule="exact"/>'
         '<w:ind w:left="100" w:rightChars="95" w:right="95" w:hangingChars="100" w:hanging="100"/>'
         '<w:jc w:val="both"/></w:pPr>'
         '<w:tblPr><w:tblBorders>'
         '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '</w:tblBorders></w:tblPr></w:style>\n')
    write_docx("v12_tableGrid", build_styles(s), build_document(tbl_style="a3"))

    # v13: minimal styles, paragraph in document.xml adds <w:jc w:val="both"/>
    write_docx("v13_jc_doc", build_styles(), build_document(extra_ppr='<w:jc w:val="both"/>'))

    # v14: minimal styles, paragraph in document.xml adds widowControl=0
    write_docx("v14_widow_doc", build_styles(), build_document(extra_ppr='<w:widowControl w:val="0"/>'))

    print(f"\nWrote phase 2 variants. Run com_measure_15076df_buildup.py to measure.")


if __name__ == "__main__":
    main()
