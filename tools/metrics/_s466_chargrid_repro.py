"""Minimal repro to derive Word's linesAndChars chars-per-line. Same docGrid as
tokumei (linePitch=292, charSpace=1453), A4, margins L/R=1080. A long CJK body
paragraph; measure where Word wraps (chars/line) vs Oxi. Isolates the grid
char-cell width formula (the open charGrid文字詰め residual)."""
import zipfile, os
OUT=r"C:\Users\ryuji\oxi-main\tools\golden-test\repros\chargrid_wrap"
os.makedirs(OUT,exist_ok=True)
CT='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'''
RELS='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DREL='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
# CJK body font MS Mincho, size varied
def styles(sz):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'''
# 60 CJK chars (numbered so we can read the wrap position), fullwidth digits + kanji
CJK="".join(f"{chr(0xFF10+(i%10))}項目" for i in range(40))  # ０項目１項目... 120 chars
def docxml(sz):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
      f'<w:p><w:r><w:t>{CJK}</w:t></w:r></w:p>'
      '<w:sectPr>'
      '<w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
      '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="851" w:footer="992" w:gutter="0"/>'
      '<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/>'
      '</w:sectPr></w:body></w:document>')
def write(path,sz):
    with zipfile.ZipFile(path,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",CT);z.writestr("_rels/.rels",RELS)
        z.writestr("word/_rels/document.xml.rels",DREL);z.writestr("word/styles.xml",styles(sz))
        z.writestr("word/document.xml",docxml(sz))
for sz,nm in [(21,"cg_mincho_10p5"),(20,"cg_mincho_10"),(18,"cg_mincho_9")]:
    p=os.path.join(OUT,nm+".docx");write(p,sz);print("wrote",nm,"sz",sz)
print("CJK len:",len(CJK))
