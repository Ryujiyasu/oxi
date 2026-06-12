# -*- coding: utf-8 -*-
"""S551 — justified (jc=both) pack-vs-stretch matrix.
J text: 国x20 、 国x24 tail (one mid-line 、) vs N text: 国x45 tail (NO punct).
need sweep via right margin; compat 15/14/absent.
Hypothesis: pack iff need <= min(fs/2, sum rem) at ANY compat; no punct ->
never pack (stretch). Verdict: L1 count + 、 advance + a stretched-char adv.
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s551_justified')
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
TEXT_J = u'国' * 20 + u'、' + u'国' * 24 + TAIL   # one compressible
TEXT_N = u'国' * 45 + TAIL                         # none
TEXT_J2 = u'国' * 12 + u'、' + u'国' * 14 + u'、' + u'国' * 17 + TAIL  # two
TEXT_J3 = u'国' * 9 + u'、' + u'国' * 10 + u'、' + u'国' * 10 + u'、' + u'国' * 12 + TAIL  # three


def build(docx, text, right_mar, compat):
    body = ('<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
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
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', settings(compat))


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    import sys
    round2 = '--round2' in sys.argv
    round3 = '--round3' in sys.argv
    if round3:
        # tighten the c15 per-count boundaries around 1.5n+2.25
        cases = (('J', TEXT_J), ('J2', TEXT_J2), ('J3', TEXT_J3))
        compats = (15,)
        needs_by = {
            'J':  ((1228, 3.8), (1234, 4.1), (1244, 4.6)),
            'J2': ((1258, 5.3), (1264, 5.6), (1274, 6.1)),
            'J3': ((1288, 6.8), (1294, 7.1), (1298, 7.3)),
        }
    elif round2:
        cases = (('J2', TEXT_J2), ('J3', TEXT_J3))
        compats = (15, 14)
        needs_by = None
    else:
        cases = (('J', TEXT_J), ('N', TEXT_N))
        compats = (15, 14, None)
        needs_by = None
    needs2 = ((1254, 5.1), (1284, 6.6), (1304, 7.6), (1334, 9.1), (1364, 10.6))
    needs1 = ((1224, 3.6), (1194, 2.1), (1254, 5.1), (1284, 6.6), (1304, 7.6))
    for tname, text in cases:
        for compat in compats:
            for right, need in (needs_by[tname] if round3 else (needs2 if round2 else needs1)):
                tag = '%s_c%s_n%g' % (tname, compat or 'abs', need)
                docx = os.path.join(OUT, 's551_%s.docx' % tag)
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
                    adv = None
                    if tname == 'J' and l1 and 20 < l1:
                        x1 = wdoc.Range(start + 20, start + 20).Information(5)
                        x2 = wdoc.Range(start + 21, start + 21).Information(5)
                        adv = x2 - x1
                    verdict = 'PACK' if l1 == 45 else ('wrap' if l1 == 44 else '?=%s' % l1)
                    print('%s: L1=%s %s toten=%s' % (tag, l1, verdict, ('%.2f' % adv) if adv else '-'))
                finally:
                    wdoc.Close(False)
finally:
    word.Quit()
