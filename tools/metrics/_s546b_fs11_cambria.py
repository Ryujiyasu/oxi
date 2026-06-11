# -*- coding: utf-8 -*-
"""S546b — autospace at fs=11 + Cambria-digit config (the gen2 regression class).
Two configs x fs {10.5, 11, 11.5}: (a) gen2-like: ascii=Cambria, ea=MS Mincho,
no kern, no compat flags, spacing line=276 lineRule=auto (docDefaults clone);
(b) MS-Mincho-explicit (sweep continuity). Measures 国国12国国 and 国1234国.
cp932-safe ASCII output."""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s546_digit')
os.makedirs(OUT, exist_ok=True)

CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')

T1 = u'国国12国国'
T4 = u'国1234国'


def styles(sz, cambria):
    if cambria:
        fonts = '<w:rFonts w:ascii="Cambria" w:eastAsia="ＭＳ 明朝" w:hAnsi="Cambria"/>'
        kern = ''
    else:
        fonts = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
        kern = '<w:kern w:val="2"/>'
    return ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>%s%s<w:sz w:val="%d"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>' % (fonts, kern, sz))


def build(docx, sz, cambria):
    body = ''.join('<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % t for t in (T1, T4))
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>'
            '<w:docGrid w:linePitch="360"/></w:sectPr>')
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>') % (body, sect)
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', styles(sz, cambria))


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for cambria in (True, False):
        for sz in (21, 22, 23):
            tag = 's546b_%s_sz%d' % ('CAM' if cambria else 'MIN', sz)
            docx = os.path.join(OUT, tag + '.docx')
            build(docx, sz, cambria)
            wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            try:
                for pi, p in enumerate(wdoc.Paragraphs):
                    rng = p.Range
                    txt = rng.Text
                    start = rng.Start
                    seq = []
                    for i in range(min(len(txt), 12)):
                        ch = txt[i]
                        if ch in ('\r', '\n', '\x07'):
                            continue
                        x = wdoc.Range(start + i, start + i).Information(5)
                        seq.append((ch, x))
                    advs = ' '.join('U+%04X=%.2f' % (ord(seq[j][0]), seq[j + 1][1] - seq[j][1])
                                    for j in range(len(seq) - 1))
                    print('%s p%d: %s' % (tag, pi, advs))
            finally:
                wdoc.Close(False)
finally:
    word.Quit()
