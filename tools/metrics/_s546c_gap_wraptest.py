# -*- coding: utf-8 -*-
"""S546c — discriminate autospace gap 2.5 vs 2.625 (=fs/4) at fs=10.5 via WRAP.
Painted positions cannot distinguish them (px granularity 0.75 > diff 0.125);
line capacity over 15 digit clusters amplifies the diff to 3.75pt > 1 char.
Para: (国12)x15 + 国x30, fs10.5 MS Mincho everywhere, no puncts (no oikomi),
no indent. Predicted L1 chars: gap=2.5 -> 52, gap=2.625 -> 51.
Also fs=11 variant: (国12)x12 + pad; gap=2.75 vs 3.0.
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
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')


def styles(sz):
    return ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
            '<w:kern w:val="2"/><w:sz w:val="%d"/></w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>' % sz)


def build(docx, sz, units, pads):
    text = (u'国12' * units) + (u'国' * pads)
    body = ('<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
            '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304"/></w:sectPr>')
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>') % (body, sect)
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', styles(sz))


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for sz, units, pads in ((21, 15, 30), (22, 12, 30), (18, 15, 30)):
        tag = 's546c_sz%d' % sz
        docx = os.path.join(OUT, tag + '.docx')
        build(docx, sz, units, pads)
        wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        try:
            p = wdoc.Paragraphs(1)
            rng = p.Range
            txt = rng.Text
            start = rng.Start
            # find first char whose Y differs from char 0 -> L1 char count
            y0 = wdoc.Range(start, start).Information(6)
            l1 = None
            for i in range(1, min(len(txt), 90)):
                ch = txt[i]
                if ch in ('\r', '\n', '\x07'):
                    continue
                y = wdoc.Range(start + i, start + i).Information(6)
                if abs(y - y0) > 0.5:
                    l1 = i
                    break
            print('%s: L1 chars = %s' % (tag, l1))
        finally:
            wdoc.Close(False)
finally:
    word.Quit()
