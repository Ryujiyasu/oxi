# -*- coding: utf-8 -*-
"""S554 — c15 justified pack boundaries at fs=12 (sz24).
fs10.5 boundaries: n1 (4.1,4.6]; n2 (6.1,6.6]; n3 (7.1,7.3].
If fs-INVARIANT (pt): same. If fs-PROPORTIONAL: ×8/7 → n1 ~(4.7,5.3].
Text: 国x18 、 国x19 + tail (38 visible); L1 natural = 456; need = 468−cap.
n2: 国x12 、 国x12 、 国x12 (38); n3: 国x9 、 国x9 、 国x9 、 国x8 (38+...).
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s554_fs12')
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
    1: u'国' * 18 + u'、' + u'国' * 19 + TAIL,
    2: u'国' * 12 + u'、' + u'国' * 12 + u'、' + u'国' * 11 + TAIL,
    3: u'国' * 9 + u'、' + u'国' * 9 + u'、' + u'国' * 8 + u'、' + u'国' * 7 + TAIL,
    4: u'国' * 7 + u'、' + u'国' * 7 + u'、' + u'国' * 7 + u'、' + u'国' * 7 + u'、' + u'国' * 5 + TAIL,
    5: u'国' * 5 + u'、' + u'国' * 6 + u'、' + u'国' * 6 + u'、' + u'国' * 6 + u'、' + u'国' * 6 + u'、' + u'国' * 3 + TAIL,
}


def build(docx, n, right_mar):
    body = ('<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % TEXTS[n])
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
    import sys
    if '--round2' in sys.argv:
        ns = (3, 4, 5)
        needs = (8.1, 8.6, 9.1, 9.6, 10.1)
    else:
        ns = (1, 2, 3)
        needs = (3.6, 4.1, 4.6, 5.1, 5.6, 6.1, 6.6, 7.1, 7.6, 8.1)
    for n in ns:
        row = []
        for need in needs:
            cap = 468.0 - need
            right = 11906 - 1304 - int(round(cap * 20))
            docx = os.path.join(OUT, 's554_n%d_%g.docx' % (n, need))
            build(docx, n, right)
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
        print('fs12 n=%d  %s' % (n, '  '.join(row)))
finally:
    word.Quit()
