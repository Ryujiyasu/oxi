# -*- coding: utf-8 -*-
"""S550 — unified oikomi derivation matrix.
Hypothesis: (1) PLAIN pull (+1 ordinary char) = demand compression up to
min(fs/2, sum halving) at ALL compat levels; (2) KINSOKU pull (+1 line-start-
prohibited char) = NO compression; compat<=14/absent -> burasagari (hang),
compat>=15-explicit -> oidashi.
Texts (fs10.5 MS Mincho, compressPunctuation, jc=left):
  P: 国x20 、 国x24 + tail  -> L1 natural 44 (ends 国), 45th = 国 (plain)
  K: 国x44 、 国x5          -> L1 natural 44 (ends 国), 45th = 、 (kinsoku)
need = 472.5 - capacity: right margin 1194 -> 2.1, 1254 -> 5.1, 1304 -> 7.6.
One compressible (、) on L1 -> budget = 5.25 (fs/2).
Verdict from L1 char count + the 、 advance (compressed vs natural-hang).
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s550_matrix')
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
          '<w:kern w:val="2"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')


def settings(compat):
    s = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
         '<w:characterSpacingControl w:val="compressPunctuation"/>')
    if compat:
        s += ('<w:compat><w:compatSetting w:name="compatibilityMode" '
              'w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat>' % compat)
    s += '</w:settings>'
    return s


TAIL = u'続きの文章がここにあります。'
TEXT_P = u'国' * 20 + u'、' + u'国' * 24 + TAIL   # 45th char = 国 (plain)
TEXT_K = u'国' * 44 + u'、' + u'国' * 5            # 45th char = 、 (kinsoku)


STYLES_NOKERN = STYLES.replace('<w:kern w:val="2"/>', '')


def build(docx, text, right_mar, compat, kern=True, grid=False):
    body = ('<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="%d" w:bottom="1134" w:left="1304"/>%s</w:sectPr>'
            % (right_mar, '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else ''))
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>') % (body, sect)
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES if kern else STYLES_NOKERN)
        z.writestr('word/settings.xml', settings(compat))


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for tname, text in (('P', TEXT_P), ('K', TEXT_K)):
        for compat in (15, 14, None):
            for right, need in ((1194, 2.1), (1254, 5.1), (1304, 7.6)):
                tag = '%s_c%s_n%g' % (tname, compat or 'abs', need)
                docx = os.path.join(OUT, 's550_%s.docx' % tag)
                build(docx, text, right, compat)
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
                    # 、 advance (P: idx20; K: idx44 if on L1)
                    ti = 20 if tname == 'P' else 44
                    adv = None
                    if l1 and ti < l1:
                        x1 = wdoc.Range(start + ti, start + ti).Information(5)
                        x2 = wdoc.Range(start + ti + 1, start + ti + 1).Information(5)
                        y2 = wdoc.Range(start + ti + 1, start + ti + 1).Information(6)
                        if abs(y2 - y0) < 0.5:
                            adv = x2 - x1
                    print('%s: L1=%s toten_adv=%s' % (tag, l1, ('%.2f' % adv) if adv else 'EOL/-'))
                finally:
                    wdoc.Close(False)
finally:
    word.Quit()
