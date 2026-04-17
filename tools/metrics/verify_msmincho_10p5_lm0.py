"""Verify Word's actual single-line gap for MS Mincho 10.5pt in LM0 (no docGrid).

Hypothesis: lm0_lineauto.json says 12.0pt but 0e7a's body uses 13.5pt gaps.
"""
import os, sys, time, zipfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = os.path.abspath("pipeline_data/msmincho_lm0_test.docx")

DOC_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
  <w:p><w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr><w:t>Line 1</w:t></w:r></w:p>
  <w:p><w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr><w:t>Line 2</w:t></w:r></w:p>
  <w:p><w:pPr><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr><w:t>Line 3</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/></w:sectPr>
</w:body>
</w:document>"""

CT = """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

RELS = """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", DOC_XML)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
time.sleep(0.3)
y1 = doc.Paragraphs(1).Range.Information(6)
y2 = doc.Paragraphs(2).Range.Information(6)
y3 = doc.Paragraphs(3).Range.Information(6)
gap12 = y2 - y1
gap23 = y3 - y2
print(f'MS Mincho 10.5pt LM0 single-line measurements:')
print(f'  P1.y = {y1}')
print(f'  P2.y = {y2}  (gap from P1: {gap12:.2f}pt)')
print(f'  P3.y = {y3}  (gap from P2: {gap23:.2f}pt)')
print(f'\n  lm0_lineauto.json says: 12.0pt')
print(f'  Word actual: {gap12:.2f}pt')
print(f'  Δ: {gap12 - 12.0:+.2f}pt')
doc.Close(SaveChanges=False)
word.Quit()
os.remove(DOCX)
