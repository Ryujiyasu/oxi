# -*- coding: utf-8 -*-
"""S548b — burasagari vs oidashi at compat15: kern-dependence test.
Para = 国x44 + 、 + 国x5, capacity 464.9 (44 fit naturally, 、 is char 45,
overflow 7.6 > fs/2 so no absorb). L1=45 -> HANG; L1=43 -> OIDASHI.
Matrix: kern {0,2} x compat {explicit15, explicit14, absent}.
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s548_oidashi')
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


def settings(compat):
    s = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
         '<w:characterSpacingControl w:val="compressPunctuation"/>')
    if compat:
        s += ('<w:compat><w:compatSetting w:name="compatibilityMode" '
              'w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat>' % compat)
    s += '</w:settings>'
    return s


def styles(kern):
    return ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
            + ('<w:kern w:val="2"/>' if kern else '') +
            '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')


def build(docx, kern, compat):
    text = u'国' * 44 + u'、' + u'国' * 5
    body = ('<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304"/></w:sectPr>')
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>') % (body, sect)
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', styles(kern))
        z.writestr('word/settings.xml', settings(compat))


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for kern in (0, 2):
        for compat in (15, 14, None):
            tag = 'k%d_c%s' % (kern, compat or 'abs')
            docx = os.path.join(OUT, 's548b_%s.docx' % tag)
            build(docx, kern, compat)
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
                verdict = 'HANG' if l1 == 45 else ('OIDASHI' if l1 == 43 else '??=%s' % l1)
                print('%s: L1=%s %s' % (tag, l1, verdict))
            finally:
                wdoc.Close(False)
finally:
    word.Quit()
