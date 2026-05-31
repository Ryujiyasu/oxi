"""Generate minimal repros isolating Cambria 11pt line-height at lineRule=auto
(line=276 => 1.15x), with/without after-spacing, to measure Word's exact line
height vs Oxi's. Each para is single-line so consecutive Y-gap = line_height(+after)."""
import zipfile, os

OUT = r"C:\Users\ryuji\oxi-main\tools\golden-test\repros\gen2_lineheight"
os.makedirs(OUT, exist_ok=True)

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

# docDefaults: sz=22 (11pt). No global spacing (set per-para to be explicit).
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>'''

def doc_xml(n, line, lineRule, after, sz=22, font="Cambria"):
    paras = []
    for i in range(n):
        paras.append(
            f'<w:p><w:pPr><w:spacing w:after="{after}" w:line="{line}" w:lineRule="{lineRule}"/>'
            f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz}"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz}"/></w:rPr>'
            f'<w:t>Line {i:02d} ABCDEFG abcdefg 0123456789</w:t></w:r></w:p>')
    body = "".join(paras)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
            f'{body}'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
            '</w:sectPr></w:body></w:document>')

def write_docx(path, document_xml):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", document_xml)

variants = {
    # name: (line, lineRule, after, sz, font)
    "cambria11_auto115_after0":   (276, "auto", 0,   22, "Cambria"),
    "cambria11_auto115_after200": (276, "auto", 200, 22, "Cambria"),
    "cambria11_auto100_after0":   (240, "auto", 0,   22, "Cambria"),   # 1.0x baseline
    "calibri11_auto115_after0":   (276, "auto", 0,   22, "Calibri"),   # cross-font
}
for name,(line,rule,after,sz,font) in variants.items():
    p = os.path.join(OUT, name + ".docx")
    write_docx(p, doc_xml(15, line, rule, after, sz, font))
    print("wrote", p)

# --- table repro: 5-row x 3-col, single-line Cambria 11pt cells ---
def table_doc(rows=5, cols=3):
    def cell(r,c):
        return ('<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
                '<w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/>'
                '<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="22"/></w:rPr></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="22"/></w:rPr>'
                f'<w:t>Cell {r}-{c}</w:t></w:r></w:p></w:tc>')
    trs=[]
    for r in range(1,rows+1):
        trs.append("<w:tr>"+"".join(cell(r,c) for c in range(1,cols+1))+"</w:tr>")
    tbl=('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
         '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
         +"".join(trs)+"</w:tbl>")
    # marker paras before and after to measure table block height
    mk=lambda t:('<w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/>'
        '<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="22"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="22"/></w:rPr><w:t>{t}</w:t></w:r></w:p>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        +mk("BEFORE_MARK")+tbl+mk("AFTER_MARK")
        +'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
        '</w:sectPr></w:body></w:document>')

p=os.path.join(OUT,"table_5x3_cambria11.docx")
write_docx(p, table_doc(5,3))
print("wrote",p)
