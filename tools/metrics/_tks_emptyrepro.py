# -*- coding: utf-8 -*-
import os, zipfile
OUT=r'C:/tmp/tks_emptyrepro.docx'
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
 '<Default Extension="xml" ContentType="application/xml"/>'
 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
 '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
RELS=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
 '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
# docDefaults matching tokyoshugyo: ascii=Century eastAsia=MS Mincho sz=21 (10.5pt)
STYLES=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
 '<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
 '<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>'
 '<w:pPrDefault/></w:docDefaults>'
 '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>') % NS
# body line (CJK so uses MS Mincho), empty para (inherit), bold heading
def body(t): return '<w:p><w:r><w:t>%s</w:t></w:r></w:p>' % t
EMPTY='<w:p/>'
HEAD='<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>（見出）</w:t></w:r></w:p>'  # （見出）bold
sect=('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
 '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992"/>'
 '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
# Structure: A, 5 empties, B, HEAD, C  (measure A->B/6=empty pitch; B->HEAD; HEAD->C)
paras = body('本文Ａ') + EMPTY*5 + body('本文Ｂ') + HEAD + body('本文Ｃ')
DOC='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s</w:body></w:document>' % (NS, paras, sect)
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',CT); z.writestr('_rels/.rels',RELS)
    z.writestr('word/_rels/document.xml.rels',DRELS)
    z.writestr('word/document.xml',DOC); z.writestr('word/styles.xml',STYLES)
print('built',OUT)
