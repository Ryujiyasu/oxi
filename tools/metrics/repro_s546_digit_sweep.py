# -*- coding: utf-8 -*-
"""S546 — digit font resolution + autospace asymmetry fs-sweep.
Questions:
  Q1: MS Mincho (hint=eastAsia) digit width per fs (expect 0.5em: 5.25/6.0/7.0)
  Q2: autospace BEFORE (CJK->digit) and AFTER (digit->CJK) per fs.
      Known @10.5: before=3.0 after=2.25 (EA digits); Century digits: after=2.0
      (8.25-6.25); before unmeasured. Models to discriminate:
        px-snap: before=ceil(fs/4 / 0.75px)*0.75, after=floor(...)?
        flat / em-scaled
  Q3: do ASCII LETTERS also route to the EA font under hint=eastAsia?
  Q4: kana<->digit boundary same autospace as ideograph?
Builds one docx per (fs, hint) combo, nogrid, jc=left, measures per-char x.
cp932-safe ASCII output.
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
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat>'
            '<w:balanceSingleByteDoubleByteWidth/><w:useFELayout/>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
            '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '</w:compat></w:settings>')


def styles(sz):
    return ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>'
            '<w:kern w:val="2"/>'
            '<w:sz w:val="%d"/></w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>' % sz)


# T1: ideograph<->digit cluster (before/after + digit width)
# T2: kana<->digit
# T3: ideograph<->LETTERS (Q3: letter font routing + DE autospace)
# T4: digit run of 4 (digit width stability, kern-independence)
T1 = u'国国12国国'        # 国国12国国
T2 = u'のの12のの'        # のの12のの
T3 = u'国国AB国国'        # 国国AB国国
T4 = u'国1234国'                  # 国1234国
# T5: halfwidth KATAKANA (no autospace boundary; pure halfwidth advance test)
T5 = u'国ｱｲｳ国'
TEXTS = [T1, T2, T3, T4, T5]


EXPLICIT_EA = os.environ.get('S546_EXPLICIT_EA', '0') == '1'


def para(text, hint):
    if EXPLICIT_EA and hint:
        rpr = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/></w:rPr>'
    elif hint:
        rpr = '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
    else:
        rpr = ''
    return ('<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
            '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (rpr, text))


def build(docx, sz, hint):
    body = ''.join(para(t, hint) for t in TEXTS)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304"/></w:sectPr>')
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>') % (body, sect)
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', styles(sz))
        z.writestr('word/settings.xml', SETTINGS)


NAMES = {0: 'T1 ideo-digit', 1: 'T2 kana-digit', 2: 'T3 ideo-ALPHA', 3: 'T4 digitrun', 4: 'T5 hwkana'}
word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for sz in (21, 24, 28, 18):  # 10.5 / 12 / 14 / 9 pt
        for hint in (True, False):
            tag = 'fs%g_%s' % (sz / 2.0, 'EA' if hint else 'CEN')
            docx = os.path.join(OUT, 's546_%s.docx' % tag)
            build(docx, sz, hint)
            wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            try:
                print('==== %s ====' % tag)
                for pi, p in enumerate(wdoc.Paragraphs):
                    rng = p.Range
                    txt = rng.Text
                    start = rng.Start
                    seq = []
                    for i in range(min(len(txt), 20)):
                        ch = txt[i]
                        if ch in ('\r', '\n', '\x07'):
                            continue
                        x = wdoc.Range(start + i, start + i).Information(5)
                        seq.append((ch, x))
                    advs = []
                    for j in range(len(seq) - 1):
                        advs.append('U+%04X=%.2f' % (ord(seq[j][0]), seq[j + 1][1] - seq[j][1]))
                    print('  %s: %s' % (NAMES[pi % 5], ' '.join(advs)))
            finally:
                wdoc.Close(False)
finally:
    word.Quit()
