"""Build 2 minimal repros to test if compressPunctuation gates the
half-width '．' (U+FF0E) advance in MS Mincho.

Repro structure: 1 paragraph in a narrow cell (similar to 15076df L12)
  - text: '１．提供を受けた匿名データの名称' (16 chars)
  - rPr: <w:rFonts hAnsi="ＭＳ 明朝" cs="ＭＳ 明朝"/><w:spacing val="-9"/><w:kern val="0"/>
  - paragraph in tcW=1968dxa cell (= 98.4pt)
  - balanceSBDB ON in settings.xml
  - VARIANT A: characterSpacingControl=compressPunctuation
  - VARIANT B: characterSpacingControl=doNotCompress (or absent)

Then COM-measure '．' advance (= position diff between '．' and '提'):
  - If A '．'=6pt AND B '．'=10.5pt → compressPunctuation gates it → fix in layout/font path
  - If both '．'=6pt → MS Mincho intrinsic → fix in compact.json metrics
  - If both '．'=10.5pt → something else drives Word's 6pt in 15076df
"""
import os
import sys
import io
import zipfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.normpath(os.path.join(REPO, "tools/metrics/_repros"))
os.makedirs(OUT_DIR, exist_ok=True)


def make_docx(out_path, compress_punc: bool):
    settings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    settings += '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
    settings += '<w:compat>\n'
    settings += '<w:balanceSingleByteDoubleByteWidth/>\n'
    settings += '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>\n'
    settings += '</w:compat>\n'
    if compress_punc:
        settings += '<w:characterSpacingControl w:val="compressPunctuation"/>\n'
    else:
        settings += '<w:characterSpacingControl w:val="doNotCompress"/>\n'
    settings += '</w:settings>\n'

    styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    styles += '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
    styles += '<w:docDefaults><w:rPrDefault><w:rPr>'
    styles += '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:cs="ＭＳ 明朝"/>'
    styles += '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/>'
    styles += '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
    styles += '</w:rPr></w:rPrDefault></w:docDefaults>\n'
    styles += '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>\n'
    styles += '</w:styles>\n'

    document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    document += '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    document += ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
    document += '<w:body>\n'
    # Single-row table with one narrow cell
    document += '<w:tbl><w:tblPr><w:tblW w:w="1968" w:type="dxa"/>'
    document += '<w:tblLayout w:type="fixed"/>'
    document += '<w:tblCellMar><w:left w:w="12" w:type="dxa"/><w:right w:w="12" w:type="dxa"/></w:tblCellMar>'
    document += '</w:tblPr>'
    document += '<w:tblGrid><w:gridCol w:w="1968"/></w:tblGrid>'
    document += '<w:tr>'
    document += '<w:tc><w:tcPr><w:tcW w:w="1968" w:type="dxa"/></w:tcPr>'
    document += '<w:p><w:pPr>'
    document += '<w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>'
    document += '<w:adjustRightInd w:val="0"/>'
    document += '<w:spacing w:line="240" w:lineRule="exact"/>'
    document += '<w:ind w:left="215" w:right="76" w:hanging="192"/>'
    document += '</w:pPr>'
    # Single run with cs=-9
    document += '<w:r><w:rPr>'
    document += '<w:rFonts w:hAnsi="ＭＳ 明朝" w:cs="ＭＳ 明朝" w:hint="eastAsia"/>'
    document += '<w:spacing w:val="-9"/><w:kern w:val="0"/><w:szCs w:val="21"/>'
    document += '</w:rPr>'
    document += '<w:t>１．提供を受けた匿名データの名称</w:t></w:r>'
    document += '</w:p></w:tc></w:tr></w:tbl>\n'
    # End body with sectPr
    document += '<w:sectPr>'
    document += '<w:pgSz w:w="11906" w:h="16838"/>'
    document += '<w:pgMar w:top="851" w:right="1134" w:bottom="567" w:left="1134"'
    document += ' w:header="851" w:footer="567" w:gutter="0"/>'
    document += '<w:docGrid w:type="lines" w:linePitch="336"/>'
    document += '</w:sectPr>\n'
    document += '</w:body></w:document>\n'

    content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    content_types += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
    content_types += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
    content_types += '<Default Extension="xml" ContentType="application/xml"/>\n'
    content_types += '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
    content_types += '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'
    content_types += '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
    content_types += '</Types>\n'

    rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
    rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n'
    rels += '</Relationships>\n'

    doc_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    doc_rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
    doc_rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
    doc_rels += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n'
    doc_rels += '</Relationships>\n'

    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/settings.xml', settings)


for cp, suffix in [(True, 'compressPunc'), (False, 'doNotCompress')]:
    out = os.path.join(OUT_DIR, f'repro_dot_15076_{suffix}.docx')
    make_docx(out, cp)
    print(f'wrote {out}')
print('\nNext: open in Word manually (or via COM) to verify they render, then run com_measure_dot_repro.py')
