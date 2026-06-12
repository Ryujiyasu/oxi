# -*- coding: utf-8 -*-
"""S548 — is the demand-oikomi light tier kern-gated like the pair rule?
Clone the S546f single-、 absorb case at kern=2 vs kern=0 (compressPunctuation
ON, compat absent, fs10.5): line of 44 chars with one mid 、, overflow set to
need {0.6, 2.1, 5.1} via right margin. If kern=0 kills the absorb, L1=43
at every need; kern=2 reproduces S546f (44 up to need 5.25).
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s548_oikomi')
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
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/></w:settings>')
TAIL = u'続きの文章がここにありますので折り返します。'


def styles(kern):
    return ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
            + ('<w:kern w:val="2"/>' if kern else '') +
            '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')


def build(docx, kern, right_mar):
    text = u'国' * 21 + u'、' + u'国' * 22 + TAIL
    body = ('<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
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
        z.writestr('word/styles.xml', styles(kern))
        z.writestr('word/settings.xml', SETTINGS)


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for kern in (True, False):
        for k in (70, 100, 160):  # need 0.60 / 2.10 / 5.10
            right = 1304 + k
            need = k / 20.0 - 2.9
            docx = os.path.join(OUT, 's548_k%d_r%d.docx' % (kern, k))
            build(docx, kern, right)
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
                print('kern=%d need=%.2f L1=%s %s' % (kern, need, l1, 'ABSORB' if l1 == 44 else 'oidashi'))
            finally:
                wdoc.Close(False)
finally:
    word.Quit()
