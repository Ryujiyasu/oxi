import zipfile, os
out='/c/tmp/sgrepro/sg_header_style.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
CT='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
RELS='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DRELS='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
# replicate ohnoikuji a4 (header) + Normal exactly
STYLES=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho" w:cs="MS Mincho"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr><w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="a4"><w:name w:val="header"/><w:basedOn w:val="a"/><w:semiHidden/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4252"/><w:tab w:val="right" w:pos="8504"/></w:tabs><w:snapToGrid w:val="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="cust"><w:name w:val="custsg0"/><w:basedOn w:val="a"/><w:pPr><w:snapToGrid w:val="0"/></w:pPr></w:style>
</w:styles>'''
def para(text, style=None, direct_sg0=False):
    ppr='<w:pPr>'
    if style: ppr+=f'<w:pStyle w:val="{style}"/>'
    if direct_sg0: ppr+='<w:snapToGrid w:val="0"/>'
    ppr+='</w:pPr>'
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
body=''
for i in range(3): body+=para(f'GRID行{i}あいうえお')
for i in range(3): body+=para(f'HEADER行{i}あいうえお', style='a4')      # built-in header style sg0
for i in range(3): body+=para(f'CUST行{i}あいうえお', style='cust')      # custom style sg0
for i in range(3): body+=para(f'DIRECT行{i}あいうえお', direct_sg0=True) # direct sg0
for i in range(3): body+=para(f'GRID2行{i}あいうえお')
sect='<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360" w:charSpace="-3426"/></w:sectPr>'
DOC=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W}"><w:body>{body}{sect}</w:body></w:document>'
with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',CT); z.writestr('_rels/.rels',RELS)
    z.writestr('word/_rels/document.xml.rels',DRELS); z.writestr('word/styles.xml',STYLES); z.writestr('word/document.xml',DOC)
print('wrote',out)
