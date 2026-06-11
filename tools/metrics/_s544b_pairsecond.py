# -*- coding: utf-8 -*-
"""S544b: does the SECOND member of a kern pair (the 、 in ）、) count toward
the -0.75 oikomi budget?
  D: 44 chars with （...）、 (） halves via pair; 、=pair-second; （=isolated),
     overflow 1.3 → needs 2 compressions. If pair-second counts: fits (44).
     If only isolated （: budget 0.75 < 1.3 → no (43).
  E: control, 2 isolated puncts （ and 、, same overflow 1.3 → expect 44.
jc=left, type=lines, MS Mincho sz=21, kern=2. ascii output.
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s543_sweep')
DOCX = os.path.join(OUT, 's544b_pairsecond.docx')

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
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
          '<w:kern w:val="2"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')

CONTENT = (11906 - 1304 * 2) / 20.0  # 464.9
OVER = 1.3

# D: 国*18 （ 国*2 ）、 国*21 = 44 chars; ） halves -> width = 44*10.5 - 5.25
TEXT_D = '国' * 18 + '（' + '国' * 2 + '）、' + '国' * 21
W_D = 44 * 10.5 - 5.25
IND_D = int(round((CONTENT - (W_D - OVER)) * 20))
# E: 国*18 （ 国*4 、 国*20 = 44 chars, no pair -> width = 462
TEXT_E = '国' * 18 + '（' + '国' * 4 + '、' + '国' * 20
W_E = 44 * 10.5
IND_E = int(round((CONTENT - (W_E - OVER)) * 20))

paras = []
for text, ind in ((TEXT_D, IND_D), (TEXT_E, IND_E)):
    paras.append('<w:p><w:pPr><w:ind w:right="%d"/><w:jc w:val="left"/></w:pPr>'
                 '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                 '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ind, text))
doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
       '<w:body>%s<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
       '<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304"/>'
       '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>') % ''.join(paras)
with zipfile.ZipFile(DOCX, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CT)
    z.writestr('_rels/.rels', RELS)
    z.writestr('word/_rels/document.xml.rels', DRELS)
    z.writestr('word/document.xml', doc)
    z.writestr('word/styles.xml', STYLES)
    z.writestr('word/settings.xml', SETTINGS)

print('IND_D=%dtw IND_E=%dtw OVER=%.2f' % (IND_D, IND_E, OVER))
word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    wdoc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    try:
        for pi, p in enumerate(wdoc.Paragraphs):
            if pi >= 2:
                break
            rng = p.Range
            start = rng.Start
            y0 = wdoc.Range(start, start).Information(6)
            n_l1 = 0
            advs = []
            prev = None
            for i in range(44):
                r = wdoc.Range(start + i, start + i)
                if r.Information(6) != y0:
                    break
                x = r.Information(5)
                if prev is not None:
                    advs.append(round(x - prev, 2))
                prev = x
                n_l1 += 1
            small = [(i, a) for i, a in enumerate(advs) if a < 10.4]
            print('para %s L1=%d (44=fits) small_advs=%s' % ('DE'[pi], n_l1, small[:8]))
    finally:
        wdoc.Close(False)
finally:
    word.Quit()
