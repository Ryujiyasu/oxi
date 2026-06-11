# -*- coding: utf-8 -*-
"""S544: does a line that already contains a PAIR compression (。） first-char
halved) still get the -0.75 demand oikomi on its other puncts? Two paras:
  A: 44 chars with 。） pair (idx21-22) + 、 (idx30), ind right tuned so the
     44th char fits ONLY if 、 compresses (overflow ~0.5 <= 0.75).
  B: control without the pair (43 chars + 、), same overflow tuning.
jc=left, type=lines, MS Mincho sz=21, kern=2 (ed025c context). ascii output.
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s543_sweep')
os.makedirs(OUT, exist_ok=True)
DOCX = os.path.join(OUT, 's544_pair_oikomi.docx')

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

# A: 44 chars, width = 44*10.5 - 5.25 (。 halved in 。）) = 456.75
TEXT_A = '国' * 21 + '。）' + '国' * 7 + '、' + '国' * 13
W_A = 44 * 10.5 - 5.25
IND_A = int(round((CONTENT - (W_A - 0.5)) * 20))
# B control: 44 chars, no pair, one 、 -> width 462.0; need ind so overflow 0.5
TEXT_B = '国' * 30 + '、' + '国' * 13
W_B = 44 * 10.5
IND_B = int(round((CONTENT - (W_B - 0.5)) * 20))
# C: pair only, NO standalone comma (can the pair itself absorb more?)
TEXT_C = '国' * 21 + '。）' + '国' * 21
W_C = 44 * 10.5 - 5.25
IND_C = IND_A

paras = []
for text, ind in ((TEXT_A, IND_A), (TEXT_B, IND_B), (TEXT_C, IND_C)):
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

print('IND_A=%dtw IND_B=%dtw' % (IND_A, IND_B))
word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    wdoc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    try:
        for pi, p in enumerate(wdoc.Paragraphs):
            if pi >= 3:
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
            print('para %s L1=%d (44=oikomi, 43=no) small_advs=%s' % ('ABC'[pi], n_l1, small[:6]))
    finally:
        wdoc.Close(False)
finally:
    word.Quit()
