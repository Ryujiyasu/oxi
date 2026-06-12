# -*- coding: utf-8 -*-
"""S547 — measure the full MS Mincho yakumono pair table under w:kern=0 vs 2.
One para per (X, Y) pair: 国XY国国. Measures advance(X) (and advance(Y)).
Pairs where advance differs from the single-char natural ONLY at kern=2 =
KERN-gated; differing at both = unconditional (S532 class).
fs=10.5 (sz21), compressPunctuation, compat absent (legacy), jc=left.
Output: c:/tmp/s547_pairs.txt (only non-natural advances).
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s547_kern')
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


def styles(kern):
    return ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
            + ('<w:kern w:val="2"/>' if kern else '') +
            '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')


CHARS = u'、。，．（）「」『』【】〔〕・：；！？［］｛｝〈〉《》'
PAIRS = [(x, y) for x in CHARS for y in CHARS]


def build(docx, kern):
    paras = []
    for x, y in PAIRS:
        paras.append('<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
                     '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                     '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (u'国' + x + y + u'国国'))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304"/></w:sectPr>')
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>') % (''.join(paras), sect)
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', styles(kern))
        z.writestr('word/settings.xml', SETTINGS)


import io
out = io.open('c:/tmp/s547_pairs.txt', 'w', encoding='utf-8')
word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for kern in (False, True):
        tag = 'kern2' if kern else 'kern0'
        docx = os.path.join(OUT, 's547_%s.docx' % tag)
        build(docx, kern)
        wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        try:
            n_nonnat = 0
            for pi, p in enumerate(wdoc.Paragraphs):
                if pi >= len(PAIRS):
                    break
                x, y = PAIRS[pi]
                rng = p.Range
                start = rng.Start
                # chars: 国 X Y 国 国
                xs = []
                for i in range(4):
                    xs.append(wdoc.Range(start + i, start + i).Information(5))
                adv_x = xs[2] - xs[1]
                adv_y = xs[3] - xs[2]
                # natural fullwidth = 10.5; flag any deviation > 0.05
                if abs(adv_x - 10.5) > 0.05 or abs(adv_y - 10.5) > 0.05:
                    n_nonnat += 1
                    out.write('%s %s%s advX=%.2f advY=%.2f\n' % (tag, x, y, adv_x, adv_y))
            out.write('%s: %d non-natural pairs\n' % (tag, n_nonnat))
            print('%s done: %d non-natural' % (tag, n_nonnat))
        finally:
            wdoc.Close(False)
finally:
    word.Quit()
    out.close()
