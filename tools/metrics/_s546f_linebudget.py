# -*- coding: utf-8 -*-
"""S546f — confirm the LINE-TOTAL demand budget (= fs/2?) vs per-punct caps.
(a) single 、 line: sweep need 3.6..6.1 -> boundary should be ~5.25 if budget
    = min(fs/2, halving) (both 5.25 at fs10.5).
(b) FOUR-punct line (、。（）): boundary ~5.25 if LINE-TOTAL fs/2; ~3.0 if
    flat 0.75/punct; >>5.25 if per-punct halving sums.
(c) fs=12 single 、: boundary 6.0 if fs/2-scaled.
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s546_digit')
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


def styles(sz):
    return ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
            '<w:kern w:val="2"/><w:sz w:val="%d"/></w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>' % sz)


def build(docx, text, right_mar, sz):
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
        z.writestr('word/styles.xml', styles(sz))
        z.writestr('word/settings.xml', SETTINGS)


TAIL = u'続きの文章がここにありますので折り返します。'
CASES = [
    # (name, text(44ch L1 target at fs10.5), sz, fullwidth_per_line, ks)
    ('a_single_toten', u'国' * 21 + u'、' + u'国' * 22 + TAIL, 21,
     [130, 140, 150, 160, 165, 170, 180]),
    ('b_four_puncts', u'国' * 10 + u'、' + u'国' * 10 + u'（' + u'国' * 5 + u'）' + u'国' * 9 + u'。' + u'国' * 7 + TAIL, 21,
     [100, 110, 120, 130, 140, 150, 160, 170]),
]

word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for name, text, sz, ks in CASES:
        for k in ks:
            right = 1304 + k
            need = k / 20.0 - 2.9
            docx = os.path.join(OUT, 's546f_%s_k%d.docx' % (name, k))
            build(docx, text, right, sz)
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
                print('%s k=%d need=%.2f L1=%s %s' % (name, k, need, l1, 'FIT' if l1 == 44 else ''))
            finally:
                wdoc.Close(False)
finally:
    word.Quit()
