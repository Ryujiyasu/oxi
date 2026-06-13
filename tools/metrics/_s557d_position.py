# -*- coding: utf-8 -*-
"""S557d — does PUNCT POSITION gate the c15 jc=both pack? Real d77a L7
(need 7.00, n=3 、、）) does NOT pack, but the S554 synthetic (n=3, evenly
spread 、) packs up to need ~8.85. Difference: L7's LAST 、 is 2 chars from
the pull point; synthetic's is ~9. Hold n=3 + need ~7.0 fixed, vary the
last punct's distance-from-line-end. If late-punct configs refuse while
spread configs pack -> position is the discriminator.

Line = 38 chars (国 filler + 、 at chosen positions), pull char = 国 (39th).
need controlled by right margin (cap = natural_38 + 国 - need ... via the
_s554 method: cap = 468 - need with the 国*38 natural = 456 baseline).
"""
import os
import sys
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s557_pos')
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

KOKU = u'国'  # 国
TOTEN = u'、'  # 、


def line_with_puncts(positions):
    """38-char line: 国 everywhere except 、 at the given 1-indexed positions."""
    chars = []
    pset = set(positions)
    for i in range(1, 39):
        chars.append(TOTEN if i in pset else KOKU)
    return u''.join(chars)


# configs: 3 puncts, vary distance of LAST punct from line end (38)
CONFIGS = {
    'spread_far':  [10, 20, 30],   # last 8 from end (S554-like)
    'mid':         [10, 20, 33],   # last 5 from end
    'late':        [10, 20, 36],   # last 2 from end (L7-like)
    'L7like':      [20, 33, 36],   # mirrors L7 positions (、@20 )@33 、@36)
}


def build(docx, line_txt, right_mar):
    body = ('<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (line_txt + TAIL))
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
needs = (6.1, 6.6, 7.1, 7.6, 8.1, 8.6, 9.1)
try:
    sys.stdout.reconfigure(encoding='utf-8')
    for name, pos in CONFIGS.items():
        line_txt = line_with_puncts(pos)
        row = []
        for need in needs:
            cap = 468.0 - need
            right = 11906 - 1304 - int(round(cap * 20))
            docx = os.path.join(OUT, 's557_%s_%g.docx' % (name, need))
            build(docx, line_txt, right)
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
        print('%-11s pos=%s  %s' % (name, pos, '  '.join(row)))
finally:
    word.Quit()
