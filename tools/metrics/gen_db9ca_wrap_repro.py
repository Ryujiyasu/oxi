"""Generate V113-V115 minimal repros for db9ca body paragraph wrap miscount.

db9ca trajectory analysis (Day 31 part 8) shows +18pt drift jumps at:
  paragraph i=20: cum drift +17pt (after page break + paragraphs 17/18/19)
  paragraph i=31: additional +18pt (paragraphs 25-30)
  paragraph i=38: cum drift +54pt (post-page boundary)

Hypothesis: Oxi wraps body paragraphs to 1 extra line vs Word, each
miscount = +18pt = linePitch=360tw=18pt.

Extracts the candidate paragraphs from db9ca:
  V113: paragraph i=18 (124 chars TNR + MS PGothic, underlined, deep indent)
  V114: paragraph i=17 (90 chars TNR + MS PGothic, underlined, mid indent)
  V115: V113 + V114 + i=19 combined

Output: tools/golden-test/repros/db9ca_wrap/DW_V113-V115.docx
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'db9ca_wrap')

# Para i=17 (90 chars, indent left=565 firstLine=353)
PARA_17 = ('<w:p>'
           '<w:pPr>'
           '<w:ind w:leftChars="269" w:left="565" w:firstLineChars="168" w:firstLine="353"/>'
           '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/>'
           '<w:bCs/><w:u w:val="single"/></w:rPr>'
           '</w:pPr>'
           '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/><w:bCs/><w:u w:val="single"/></w:rPr>'
           '<w:t xml:space="preserve">Source: Agency D website (URL of the relevant page) PDL1.0 (The License Original page URL)</w:t>'
           '</w:r>'
           '</w:p>')

# Para i=18 (124 chars, indent left=918, no firstLine)
PARA_18 = ('<w:p>'
           '<w:pPr>'
           '<w:ind w:leftChars="437" w:left="918"/>'
           '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/>'
           '<w:bCs/><w:u w:val="single"/></w:rPr>'
           '</w:pPr>'
           '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/><w:bCs/><w:u w:val="single"/></w:rPr>'
           '<w:t xml:space="preserve">Source: XX Survey (Agency D) (URL of the relevant page) PDL1.0 (The License Original page URL) (accessed on year/month/day) </w:t>'
           '</w:r>'
           '</w:p>')

# Para i=19 truncated (just first sentence, 100 chars)
PARA_19 = ('<w:p>'
           '<w:pPr>'
           '<w:ind w:leftChars="202" w:left="565" w:hangingChars="67" w:hanging="141"/>'
           '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/></w:rPr>'
           '</w:pPr>'
           '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/></w:rPr>'
           '<w:t xml:space="preserve">b. If the user has edited "This Content" for use, you must include a statement expressing</w:t>'
           '</w:r>'
           '</w:p>')

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ Ｐゴシック" w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="ＭＳ Ｐゴシック"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""


def doc_xml(body: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body}
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/>
<w:docGrid w:type="linesAndChars" w:linePitch="360"/>
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
    write_docx('DW_V113_para18_only', doc_xml(PARA_18))
    write_docx('DW_V114_para17_only', doc_xml(PARA_17))
    write_docx('DW_V115_combined_17_18_19', doc_xml(PARA_17 + PARA_18 + PARA_19))
    print('Done.')


if __name__ == '__main__':
    main()
