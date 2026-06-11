# -*- coding: utf-8 -*-
"""S542 — minimal repros for the S541 width hypotheses (7f272a p1 class):
  H1: mid-line COMMA+OPENPAREN pair: Word halves the comma (5.25) even at jc=left
  H2: jc=left mid-line punct (fullwidth period / parens / comma before CJK)
      gets -0.75 under docGrid=NONE (S492's zero-compression was type=lines)
  H3: halfwidth digit system: digit 5.25, CJK->digit autospace 3.0 (adv 13.5),
      digit->CJK 2.25 (adv 7.5)
Builds TWO docx (docGrid none / type=lines), same paragraphs at jc=left and
jc=both, MS Mincho sz=21 (10.5pt = the 7f272a context), compressPunctuation,
compat 15. COM-measures per-char advances. cp932-safe ASCII output.
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s542_width')
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
_F = os.environ.get('S542_FLAGS', 'bal,fel,otf,c14').split(',')
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat>'
            + ('<w:balanceSingleByteDoubleByteWidth/>' if 'bal' in _F else '')
            + ('<w:useFELayout/>' if 'fel' in _F else '')
            + ('<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%s"/>'
               % ('14' if 'c14' in _F else '15'))
            + ('<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>' if 'otf' in _F else '')
            + '</w:compat></w:settings>')
KERN = os.environ.get('S542_KERN', '1') == '1'
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>'
          + ('<w:kern w:val="2"/>' if KERN else '') +
          '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')

# paragraph texts (short single lines, mid-line targets):
# T1 H1: comma+openparen pair mid-line:      国国は、（　）の国国
# T2 H2: period/comma/parens before CJK:     ３．国国（国）国、国国
# T3 H3: digit cluster:                      第17条国国国
T1 = '国国は、（　）の国国'
# T2 = the VERBATIM 7f272a ３． para (wraps; L1 natural 467.25 vs measure 464.9
# → fits only with -0.75 x3 punct compression = the oikomi-at-jc=left test)
T2 = '３．第２条第３項（第17条第３項）に掲げる添付書類のうち、当該変更に伴いその内容が変更されるものを添付すること。'
T3 = '第17条国つ、れ（こ）て'
T2_IND = '<w:spacing w:line="340" w:lineRule="exact"/><w:ind w:left="283" w:hangingChars="135" w:hanging="283"/>'


RUN_RPR = ('<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/></w:rPr>'
           if os.environ.get('S542_EAFONT', '1') == '1' else '')


def para(text, jc):
    extra = T2_IND if text is T2 else ''
    return ('<w:p><w:pPr>%s<w:jc w:val="%s"/></w:pPr>'
            '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (extra, jc, RUN_RPR, text))


def build(docx, grid):
    body = ''
    for jc in ('left', 'both'):
        for t in (T1, T2, T3):
            body += para(t, jc)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304"/>%s</w:sectPr>'
            % ('<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else ''))
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>') % (body, sect)
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)


NAMES = {0: 'T1 comma+paren', 1: 'T2 period/paren', 2: 'T3 digit'}
word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for grid in (False, True):
        docx = os.path.join(OUT, 's542_%s.docx' % ('lines' if grid else 'nogrid'))
        build(docx, grid)
        wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        try:
            print('==== %s ====' % os.path.basename(docx))
            for pi, p in enumerate(wdoc.Paragraphs):
                rng = p.Range
                txt = rng.Text
                start = rng.Start
                seq = []
                for i in range(min(len(txt), 50)):
                    ch = txt[i]
                    if ch in ('\r', '\n', '\x07'):
                        continue
                    x = wdoc.Range(start + i, start + i).Information(5)
                    seq.append((ch, x))
                jc = 'left' if pi < 3 else 'both'
                print('para %d (%s, %s):' % (pi, jc, NAMES[pi % 3]))
                for j in range(len(seq) - 1):
                    adv = round(seq[j + 1][1] - seq[j][1], 2)
                    print('   U+%04X adv=%.2f' % (ord(seq[j][0]), adv))
        finally:
            wdoc.Close(False)
finally:
    word.Quit()
