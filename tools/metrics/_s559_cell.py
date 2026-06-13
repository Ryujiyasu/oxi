# -*- coding: utf-8 -*-
"""S559 — minimal 1-cell-table repro of 3a4f para 2234 (⑦) to isolate WHY Oxi
over-packs it to 1 line where Word wraps to 2. Variants mutate one var at a
time: base (firstLine=210,left=0,tcW=8458,gridCol=8244) | hanging (left=420
hanging=210, the sibling style) | tcweqgrid (tcW=8244) | nofl (firstLine=0).
Renders Oxi (GDI --dump-layout, counts ⑦ lines) and Word (COM, line count).
Usage: python _s559_cell.py [variant]
"""
import os
import sys
import zipfile

OUT = os.path.abspath('tools/golden-test/repros/s559_cell')
os.makedirs(OUT, exist_ok=True)

TEXT = u'⑦　常に整理整頓に努め、通路、避難口又は消火設備のある所に物品を置かないこと。'

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
# Normal a (jc=both, kern, sz21), a7 List Paragraph basedOn a
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>'
          '</w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>'
          '<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="a7"><w:name w:val="List Paragraph"/><w:basedOn w:val="a"/>'
          '<w:pPr><w:ind w:leftChars="400" w:left="840"/></w:pPr></w:style></w:styles>')
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def build(docx, ind_xml, tcw, gridcol):
    para = ('<w:p><w:pPr><w:pStyle w:val="a7"/>%s'
            '<w:rPr><w:szCs w:val="21"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ind_xml, TEXT))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="250" w:type="dxa"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
           '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
           '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
           '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>' % (gridcol, tcw, para))
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
    'base':      ('<w:ind w:leftChars="0" w:left="0" w:firstLineChars="100" w:firstLine="210"/>', 8458, 8244),
    'hanging':   ('<w:ind w:leftChars="100" w:left="420" w:hangingChars="100" w:hanging="210"/>', 8458, 8244),
    'tcweqgrid': ('<w:ind w:leftChars="0" w:left="0" w:firstLineChars="100" w:firstLine="210"/>', 8244, 8244),
    'nofl':      ('<w:ind w:leftChars="0" w:left="0"/>', 8458, 8244),
}

variant = sys.argv[1] if len(sys.argv) > 1 else 'base'
ind, tcw, gridcol = VARIANTS[variant]
docx = os.path.join(OUT, 's559_%s.docx' % variant)
build(docx, ind, tcw, gridcol)

# Oxi via GDI dump
import subprocess
GDI = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
dump = os.path.join(OUT, 's559_%s.json' % variant)
subprocess.run([GDI, docx, os.path.join(OUT, 's559_%s' % variant), '--dump-layout=' + dump],
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
import json
d = json.load(open(dump, encoding='utf-8'))
from collections import defaultdict
ys = defaultdict(list)
for pg in d['pages']:
    for e in pg['elements']:
        if e['type'] == 'text' and e['text'].strip():
            ys[round(e['y'], 1)].append(e)
sys.stdout.reconfigure(encoding='utf-8')
oxi_lines = []
for y in sorted(ys):
    es = sorted(ys[y], key=lambda e: e['x'])
    oxi_lines.append((round(es[0]['x'], 1), round(es[-1]['x'] + es[-1]['w'], 1), ''.join(e['text'] for e in es)))
print('=== %s ===' % variant)
print('OXI lines: %d' % len(oxi_lines))
for x0, x1, t in oxi_lines:
    print('  x0=%.1f x1=%.1f n=%d %s' % (x0, x1, len(t), t))

# Word via COM
try:
    import win32com.client as w32
    word = w32.DispatchEx('Word.Application')
    word.Visible = False
    try:
        wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        try:
            pr = wdoc.Paragraphs(1).Range
            s = pr.Start
            txt = pr.Text
            y0 = None
            n = 0
            counts = []
            for i in range(min(len(txt), 60)):
                ch = txt[i]
                if ch in ('\r', '\n', '\x07'):
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
            print('WORD lines: %d  counts=%s' % (len(counts), counts))
        finally:
            wdoc.Close(False)
    finally:
        word.Quit()
except Exception as e:
    print('WORD COM failed:', e)
