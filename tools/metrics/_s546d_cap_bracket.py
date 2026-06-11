# -*- coding: utf-8 -*-
"""S546d — bracket the S543 light-tier per-punct cap (0.75 vs 0.8 vs more).
The 卸売 real-doc line: overflow 2.35, 3 puncts (（）、), Word FITS -> cap >= 0.7833.
Clone its shape (42 fullwidth incl （）、 + 国12 cluster + firstLine 210 indent =
natural 467.25, base capacity 464.9) and vary the RIGHT margin to step the
overflow: right=1304 -> 2.35, 1301 -> 2.20, 1307 -> 2.50, 1310 -> 2.65,
1313 -> 2.80. Fit pattern pins the cap in 0.05pt steps.
L1 fits 44 chars iff overflow <= 3*cap.
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
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
          '<w:kern w:val="2"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')

# verbatim 卸売 para (44-char L1 shape + continuation so the para wraps)
TEXT = (u'卸売市場法第６条第１項（第14条において準用する同法第６条第１項）の規定により、中央卸売'
        u'市場（地方卸売市場）に係る認定事項の変更について認定を受けたいので、次のとおり申請します。')


def build(docx, right_mar):
    body = ('<w:p><w:pPr><w:spacing w:line="340" w:lineRule="exact"/>'
            '<w:ind w:firstLineChars="100" w:firstLine="210"/><w:jc w:val="left"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % TEXT)
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
try:
    for right in (1322, 1334, 1346, 1358, 1370, 1382, 1394, 1406):
        ovf = 2.35 + (right - 1304) / 20.0
        docx = os.path.join(OUT, 's546d_r%d.docx' % right)
        build(docx, right)
        wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        try:
            p = wdoc.Paragraphs(1)
            rng = p.Range
            txt = rng.Text
            start = rng.Start
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
            need = ovf / 3.0
            print('right=%d overflow=%.2f L1=%s (fit44 needs cap>=%.4f)' % (right, ovf, l1, need))
        finally:
            wdoc.Close(False)
finally:
    word.Quit()
