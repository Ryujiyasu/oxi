# -*- coding: utf-8 -*-
"""S559 wheel-(b) — resolve the jc-vs-left CONFOUND for the default-cellMar
reserve. In 3a4f, ⑦ (jc=both, left=0) reserves Word's default 108tw cellMar but
the structurally-identical p19 (jc=left, left=459) does not — so the gate keyed on
"justified". But ⑦ is ALSO left=0 and p19 left>0, so jc and left=0 are confounded.

This builds a 2x2 matrix (jc∈{both,left} × left∈{0,459}) in a SINGLE-cell
auto-width table whose autofit is FORCED to squeeze (tcW 8458 > gridCol) by a long
spacer paragraph that caps the column at the page-available width. For each cell we
COM-measure Word's line-1 wrap (char count) and compare to the no-reserve vs
reserve predictions, so the truth table reads:
  - if BOTH jc=both variants reserve and BOTH jc=left don't → JC is the driver
  - if BOTH left=0 variants reserve and BOTH left=459 don't → LEFT==0 is the driver
"""
import json
import os
import subprocess
import sys
import zipfile
from collections import defaultdict

OUT = os.path.abspath('tools/golden-test/repros/s559_jcleft')
os.makedirs(OUT, exist_ok=True)
sys.stdout.reconfigure(encoding='utf-8')

# the ⑦ text (39 fullwidth chars) — fills near a body-width line at fs10.5
TEXT = u'常に整理整頓に努め、通路、避難口又は消火設備のある所に物品を置かないことを徹底する。'
# spacer: a long unbreakable-ish full-width run to force autofit to body width
SPACER = u'あ' * 70

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
          '<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>'
          '</w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>'
          '<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:style></w:styles>')
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def para(jc, left, first):
    # explicit jc on the test paragraph; left/firstLine indents
    if left == 0:
        ind = '<w:ind w:leftChars="0" w:left="0" w:firstLineChars="100" w:firstLine="%d"/>' % first
    else:
        ind = '<w:ind w:left="%d" w:firstLineChars="84" w:firstLine="%d"/>' % (left, first)
    jcx = '<w:jc w:val="%s"/>' % jc
    return ('<w:p><w:pPr>%s%s<w:rPr><w:szCs w:val="21"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (jcx, ind, TEXT))


def spacer_para():
    return ('<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % SPACER)


def build(docx, jc, left, first, tcw=8458, gridcol=8244):
    cellcontent = spacer_para() + para(jc, left, first)
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="250" w:type="dxa"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
           '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>'
           % (gridcol, tcw, cellcontent))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>') % (tbl, sect)
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)


VARIANTS = {
    'just_l0':  ('both', 0,   210),  # = ⑦ config
    'left_l0':  ('left', 0,   210),
    'just_lN':  ('both', 459, 176),
    'left_lN':  ('left', 459, 176),  # = p19 config
}

GDI = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')


def oxi_test_lines(docx, tag):
    dump = os.path.join(OUT, 's559jl_%s.json' % tag)
    subprocess.run([GDI, docx, os.path.join(OUT, 's559jl_%s' % tag), '--dump-layout=' + dump],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    d = json.load(open(dump, encoding='utf-8'))
    ys = defaultdict(list)
    for pg in d['pages']:
        for e in pg['elements']:
            if e['type'] == 'text' and e['text'].strip() and u'あ' not in e['text']:
                ys[round(e['y'], 1)].append(e)
    lines = []
    for y in sorted(ys):
        es = sorted(ys[y], key=lambda e: e['x'])
        lines.append(''.join(e['text'] for e in es))
    return lines


def word_test_wrap(docx):
    import win32com.client as w32
    word = w32.DispatchEx('Word.Application')
    word.Visible = False
    try:
        wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        try:
            # the TEST paragraph is the one containing TEXT[:6]
            anchor = TEXT[:6]
            rng = wdoc.Content
            f = rng.Find
            f.Text = anchor
            if not f.Execute():
                return None
            p = rng.Paragraphs(1).Range
            s, e = p.Start, p.End
            t = p.Text
            counts = []
            y0 = None
            n = 0
            for i in range(min(e - s, 80)):
                ch = t[i] if i < len(t) else ''
                if ch in ('\r', '\n', '\x07', '\x0b'):
                    continue
                y = wdoc.Range(s + i, s + i).Information(6)
                if y0 is None:
                    y0 = y
                if abs(y - y0) > 0.5:
                    counts.append(n)
                    n = 0
                    y0 = y
                n += 1
            counts.append(n)
            return counts
        finally:
            wdoc.Close(False)
    finally:
        word.Quit()


def main():
    print('TEXT len = %d chars' % len(TEXT))
    results = {}
    for tag, (jc, left, first) in VARIANTS.items():
        docx = os.path.join(OUT, 's559jl_%s.docx' % tag)
        build(docx, jc, left, first)
        oxi = oxi_test_lines(docx, tag)
        try:
            word = word_test_wrap(docx)
        except Exception as ex:
            word = 'COM FAIL: %s' % ex
        results[tag] = (jc, left, oxi, word)
        print('\n=== %s (jc=%s left=%d) ===' % (tag, jc, left))
        print('  OXI test-para lines=%d  L1=%r' % (len(oxi), oxi[0] if oxi else ''))
        print('  WORD test-para line counts = %s' % (word,))

    print('\n===== TRUTH TABLE (Word line-1 char count) =====')
    for tag in ('just_l0', 'left_l0', 'just_lN', 'left_lN'):
        jc, left, oxi, word = results[tag]
        l1 = word[0] if isinstance(word, list) and word else '?'
        print('  %-9s jc=%-4s left=%-3d  Word L1=%s  (lines=%s)'
              % (tag, jc, left, l1, len(word) if isinstance(word, list) else word))
    print('\nRead: compare just_l0 vs left_l0 (jc effect at left=0) and '
          'just_l0 vs just_lN (left effect at jc=both).')


if __name__ == '__main__':
    main()
