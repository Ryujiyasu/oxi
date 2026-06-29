# -*- coding: utf-8 -*-
# Controlled repro: how does Word render line=350 lineRule=atLeast EMPTY paragraphs
# in a type=lines linePitch=360 (18pt) grid? Build a minimal docx, COM-measure per-para
# Information(6) Y gaps. Also measures the 注-style before/after=120 space.
import os, sys, zipfile, shutil
sys.stdout.reconfigure(encoding="utf-8")
OUT=r"C:/tmp/atleast_repro.docx"
CT='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'''
RELS='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
def para(text, spacing="", rpr='<w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'):
    sp=f'<w:spacing {spacing}/>' if spacing else ''
    run=f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>' if text else ''
    return f'<w:p><w:pPr>{sp}{rpr}</w:pPr>{run}</w:p>'
body=[]
body.append(para("CONTENT-A 本文Ａ"))
body.append(para("CONTENT-B 本文Ｂ"))
# 注-style para with before/after=120 (6pt)
body.append(para("（注）注書きテスト", spacing='w:before="120" w:after="120"'))
# 4 empty paras line=350 atLeast
for _ in range(4): body.append(para("", spacing='w:line="350" w:lineRule="atLeast"'))
body.append(para("MARK-ATLEAST マーク１", spacing='w:line="350" w:lineRule="atLeast"'))
# 4 empty paras with NO spacing (default grid)
for _ in range(4): body.append(para(""))
body.append(para("MARK-DEFAULT マーク２"))
# 4 empty paras line=350 EXACT
for _ in range(4): body.append(para("", spacing='w:line="350" w:lineRule="exact"'))
body.append(para("MARK-EXACT マーク３", spacing='w:line="350" w:lineRule="exact"'))
sect='<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1701" w:right="1701" w:bottom="1701" w:left="1701"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'
DOC=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{''.join(body)}{sect}</w:body></w:document>'''
if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',CT); z.writestr('_rels/.rels',RELS); z.writestr('word/document.xml',DOC)
print("built",OUT)
# COM measure
import win32com.client as w
app=w.Dispatch('Word.Application'); app.Visible=False
doc=app.Documents.Open(os.path.abspath(OUT))
prev=None
for i in range(1, doc.Paragraphs.Count+1):
    rng=doc.Paragraphs(i).Range
    y=doc.Range(rng.Start,rng.Start).Information(6)  # wdVerticalPositionRelativeToPage
    txt=rng.Text.strip()[:18]
    gap=(y-prev) if prev is not None else 0
    print(f"  p{i:>2} y={y:7.2f} gap={gap:6.2f}  {txt!r}")
    prev=y
doc.Close(False); app.Quit()
