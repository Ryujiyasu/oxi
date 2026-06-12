# -*- coding: utf-8 -*-
"""S555 — does a LATIN run (+autoSpaceDN) on the line kill the c15 justified
pack? (the last untested d77a-L0 feature). fs12, c15, jc=both.
L1 = 国x10 "1.0" 国x8 、 国x16 (38 chars: 34 fw + 3 hw + 、);
natural = 34x12 + 3x6 + 2x3.0(autospace) + ... = 444; 39th char need = 456−cap.
n=1 (、). Control boundaries (no latin, fs12 n1): pack ≤5.1, refuse ≥5.6.
If latin kills the pack: wrap at ALL needs.
Variant L2: n=2 (、x2) same question (control boundary 7.85).
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s555_latin')
os.makedirs(OUT, exist_ok=True)

CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
          '<w:kern w:val="2"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
TAIL = u'続きの文章がここにあります。'
TEXTS = {
    'L1': u'国' * 10 + u'1.0' + u'国' * 8 + u'、' + u'国' * 16 + TAIL,
    'L2': u'国' * 8 + u'1.0' + u'国' * 6 + u'、' + u'国' * 8 + u'、' + u'国' * 9 + TAIL,
}
# natural L1(38ch incl latin3): 34*12 + 3*6 + 6(autospace 3+3) + ... compute live:
# fw count: L1: 38-3(latin)=35? chars: 10+3+8+1+16 = 38: fw = 10+8+16=34 + 、1 = 35 fw
# natural = 35*12 + 3*6 + 6 = 420+18+6 = 444 ✓ ; L2: 8+3+6+1+8+1+9 = 36?? adjust
# L2 chars: 8+3+6+1+8+1+9 = 36 + need 2 more fw: use 国x11 tail-side
TEXTS['L2'] = u'国' * 8 + u'1.0' + u'国' * 6 + u'、' + u'国' * 8 + u'、' + u'国' * 11 + TAIL
# L2: 8+3+6+1+8+1+11 = 38 ✓ fw = 8+6+8+11+2(、、)=35 → natural = 444 ✓ same


def build(docx, text, right_mar):
    body = ('<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="%d" w:bottom="1134" w:left="1304"/></w:sectPr>' % right_mar)
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>') % (body, sect)
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for tname in ('L1', 'L2'):
        row = []
        for need in (2.1, 3.6, 4.6, 5.1, 6.1, 7.1, 7.6, 8.1):
            cap = 456.0 - need
            right = 11906 - 1304 - int(round(cap * 20))
            docx = os.path.join(OUT, 's555_%s_n%g.docx' % (tname, need))
            build(docx, TEXTS[tname], right)
            wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            try:
                pr = wdoc.Paragraphs(1).Range
                start = pr.Start
                txt = pr.Text
                y0 = wdoc.Range(start, start).Information(6)
                l1 = None
                for i in range(1, min(len(txt), 60)):
                    ch = txt[i]
                    if ch in ('\r', '\n', '\x07'):
                        continue
                    y = wdoc.Range(start + i, start + i).Information(6)
                    if abs(y - y0) > 0.5:
                        l1 = i
                        break
                row.append('%g:%s' % (need, 'P' if l1 == 39 else ('w' if l1 == 38 else '?%s' % l1)))
            finally:
                wdoc.Close(False)
        print('%s  %s' % (tname, '  '.join(row)))
finally:
    word.Quit()
