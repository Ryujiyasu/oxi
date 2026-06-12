# -*- coding: utf-8 -*-
"""S547b — pair-halving gate matrix: (kern 0/2) x (compressPunctuation on/off)
x (compat absent/15). Probe pairs: 、（ (rule A) and （「 (rule B), fs10.5.
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


def settings(csc, compat15):
    s = '<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    if csc:
        s += '<w:characterSpacingControl w:val="compressPunctuation"/>'
    else:
        s += '<w:characterSpacingControl w:val="doNotCompress"/>'
    if compat15:
        s += ('<w:compat><w:compatSetting w:name="compatibilityMode" '
              'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>')
    s += '</w:settings>'
    return s


def styles(kern):
    return ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
            + ('<w:kern w:val="2"/>' if kern else '') +
            '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')


def build(docx, kern, csc, compat15):
    paras = []
    for t in (u'国、（国国', u'国（「国国'):
        paras.append('<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
                     '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                     '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % t)
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
        z.writestr('word/settings.xml', settings(csc, compat15))


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for kern in (False, True):
        for csc in (False, True):
            for c15 in (False, True):
                tag = 'k%d_c%d_v%d' % (kern, csc, c15)
                docx = os.path.join(OUT, 's547b_%s.docx' % tag)
                build(docx, kern, csc, c15)
                wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
                try:
                    res = []
                    for p in list(wdoc.Paragraphs)[:2]:
                        start = p.Range.Start
                        xs = [wdoc.Range(start + i, start + i).Information(5) for i in range(4)]
                        res.append('%.2f/%.2f' % (xs[2] - xs[1], xs[3] - xs[2]))
                    print('%s: toten-paren=%s paren-kagi=%s' % (tag, res[0], res[1]))
                finally:
                    wdoc.Close(False)
finally:
    word.Quit()
