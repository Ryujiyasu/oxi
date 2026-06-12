# -*- coding: utf-8 -*-
"""S553b — does the c15 justified 12-quanta line budget count ALREADY-CONSUMED
pair/light compression? Text P (pair): 国x10 」（ 国x8 、 国x23 + tail —
the 」 pair-halves (−5.25 = 7q at fs10.5) at natural layout; remaining
budget = 12q − 7q = 5q?? (or per the consumed hypothesis with depth caps).
Control C: same without the pair (「 placed non-adjacent).
Prediction (consumed-budget): P refuses needs that C packs.
fs10.5, kern=2, compat15, jc=both.
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s553_consumed')
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
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
TAIL = u'続きの文章がここにあります。'
# P: 44 visible chars incl 」（ pair + virgin 、; natural L1 = 42*10.5+5.25+10.5 = 456.75
TEXT_P = u'国' * 10 + u'」（' + u'国' * 8 + u'、' + u'国' * 23 + TAIL
# C: same chars but pair broken (」 not before （): 国」国（... no adjacency
TEXT_C = u'国' * 10 + u'」' + u'国' * 4 + u'（' + u'国' * 4 + u'、' + u'国' * 23 + TAIL

word = w32.DispatchEx('Word.Application')
word.Visible = False


def build(docx, text, right_mar):
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
        z.writestr('word/settings.xml', SETTINGS)


try:
    # P natural L1(44ch) = 456.75 (pair halved); 45th char need = 467.25-cap
    # C natural L1(44ch) = 462.0 (no halving)?? C has 45 visible? keep same char count:
    # C text = 44 visible too (10+1+4+1+4+1+23 = 44)
    for tname, text, nat45 in (('P', TEXT_P, 467.25), ('C', TEXT_C, 472.5)):
        row = []
        for need in (1.4, 2.1, 3.6, 4.6, 6.1):
            cap = nat45 - need
            right = 11906 - 1304 - int(round(cap * 20))
            docx = os.path.join(OUT, 's553b_%s_n%g.docx' % (tname, need))
            build(docx, text, right)
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
                row.append('%g:%s' % (need, 'P' if l1 == 45 else ('w' if l1 == 44 else '?%s' % l1)))
            finally:
                wdoc.Close(False)
        print('%s  %s' % (tname, '  '.join(row)))
finally:
    word.Quit()
