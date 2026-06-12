# -*- coding: utf-8 -*-
"""S550b — why does 3a4f (compat15-explicit) plain-oikomi when the synthetic
compat15 doesn't? Knobs: kern {2, none} x docGrid {none, lines360}, fixed
P-text plain pull at need 5.1 (and 2.1). Expect to isolate the enabling knob."""
import importlib.util
import os
import sys

spec = importlib.util.spec_from_file_location(
    'm', os.path.join(os.path.dirname(__file__), '_s550_matrix.py'))
# can't import (it runs at import); inline instead
import zipfile

import win32com.client as w32

sys.path.insert(0, os.path.dirname(__file__))

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


def styles(kern):
    return ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
            + ('<w:kern w:val="2"/>' if kern else '') +
            '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')


def settings_xml(flags):
    return ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            + ('<w:doNotExpandShiftReturn/>' if 'dnesr' in flags else '')
            + '<w:compat>'
            + ('<w:balanceSingleByteDoubleByteWidth/>' if 'bal' in flags else '')
            + ('<w:useFELayout/>' if 'fel' in flags else '')
            + '<w:compatSetting w:name="compatibilityMode" '
              'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            + ('<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>' if 'otf' in flags else '')
            + '</w:compat></w:settings>')

TAIL = u'続きの文章がここにあります。'
TEXT_P = u'国' * 20 + u'、' + u'国' * 24 + TAIL


def build(docx, right_mar, kern, grid, flags=()):
    body = ('<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % TEXT_P)
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
        z.writestr('word/styles.xml', styles(kern))
        z.writestr('word/settings.xml', settings_xml(flags))


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for flags in (('fel',), ('bal',), ('dnesr',), ('fel', 'bal', 'otf', 'dnesr')):
        for right, need in ((1254, 5.1),):
            kern, grid = False, True
            tag = '%s_n%g' % ('+'.join(flags), need)
            docx = os.path.join(OUT, 's550b_%s.docx' % tag.replace('+', '_'))
            build(docx, right, kern, grid, flags)
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
                if l1 and 20 < l1:
                    x1 = wdoc.Range(start + 20, start + 20).Information(5)
                    x2 = wdoc.Range(start + 21, start + 21).Information(5)
                    adv = x2 - x1
                print('c15 %s: L1=%s toten=%s' % (tag, l1, ('%.2f' % adv) if adv else '-'))
            finally:
                wdoc.Close(False)
finally:
    word.Quit()
