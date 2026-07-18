# -*- coding: utf-8 -*-
"""Exact-spacer ladder: measure Word's cbot for the 002c1ffa footer shape.

Host: real styles.xml + minimal settings(compat14) + footer4.xml verbatim as
the default footer. Body: K filler paras (line=240 auto, TNR 11) + ONE
line=X exact empty spacer + a TARGET line. Sweep X in 2tw steps; the page-1
capacity flip pins cbot to 0.1pt:
  target stays on p1  iff  K*L + X/20 + L <= cbot - top(70.9)
"""
import os, sys, zipfile, re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_pbdrpunct_gen as P

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
      "pipeline_data", "_pb_ftrstack")
z = zipfile.ZipFile(P.SRC)
FOOTER = z.read('word/footer4.xml').decode('utf-8')

CT = P.CT.replace('</Types>',
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
DOCRELS = P.DOCRELS.replace('</Relationships>',
    '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')

# body filler: TNR 11 line=240 auto paras (hhea 12.649 each)
FILL = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        '<w:rPr><w:sz w:val="22"/></w:rPr></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>Filler line {i} lorem ipsum dolor.</w:t></w:r></w:p>')
SPACER = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{x}" w:lineRule="exact"/>'
          '<w:rPr><w:sz w:val="22"/></w:rPr></w:pPr></w:p>')
TARGET = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
          '<w:rPr><w:sz w:val="22"/></w:rPr></w:pPr>'
          '<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>TARGETLINE</w:t></w:r></w:p>')

K = 30  # 30 fillers = 30*12.649 = 379.5

def build(x_tw, path):
    body = ''.join(FILL.replace('{i}', str(i)) for i in range(K))
    body += SPACER.replace('{x}', str(x_tw))
    body += TARGET
    body += ('<w:sectPr><w:ftr/><w:footerReference w:type="default" r:id="rIdF"/>'
             .replace('<w:ftr/>', '') +
             '<w:pgSz w:w="11907" w:h="16839"/>'
             '<w:pgMar w:top="1418" w:right="2410" w:bottom="4252" '
             'w:left="2410" w:header="720" w:footer="3402" w:gutter="0"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {P.W_NS}><w:body>{body}</w:body></w:document>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as o:
        o.writestr('[Content_Types].xml', CT)
        o.writestr('_rels/.rels', P.RELS)
        o.writestr('word/_rels/document.xml.rels', DOCRELS)
        o.writestr('word/document.xml', doc)
        o.writestr('word/styles.xml', P.STYLES)
        o.writestr('word/settings.xml', P.SETTINGS)
        o.writestr('word/footer1.xml', FOOTER)

# capacity used before target = 70.9(top) + K*12.649 + X/20 + target line 12.649
# target on p1 iff 70.9 + K*12.649 + X/20 + 12.649 <= cbot
# cbot candidates: Oxi 613.05, Word-model 607.85 (or 608.2)
# X/20 = cbot - 70.9 - 31*12.649 -> for cbot 607.85: X = (607.85-70.9-392.1)*20 = 2896
# sweep X 2700..3100 step 10tw (0.5pt), then refine
CASES = list(range(2700, 3101, 10))

def gen():
    os.makedirs(OUT, exist_ok=True)
    for x in CASES:
        build(x, os.path.join(OUT, f'fl_{x}.docx'))
    print('generated', len(CASES))

def measure(cases=None):
    import win32com.client, fitz
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for x in (cases or CASES):
            p = os.path.join(OUT, f'fl_{x}.docx')
            pdf = p.replace('.docx', '.pdf')
            if not os.path.exists(pdf):
                doc = word.Documents.Open(p, ReadOnly=True)
                doc.ExportAsFixedFormat(pdf, 17)
                doc.Close(False)
            d = fitz.open(pdf)
            # find TARGETLINE page
            tp = None
            for pn in range(d.page_count):
                if 'TARGETLINE' in d[pn].get_text():
                    tp = pn + 1
                    break
            res[x] = (tp, d.page_count)
            d.close()
    finally:
        word.Quit()
    seq = sorted(res.items())
    flips = [(a, ra, b, rb) for (a, ra), (b, rb) in zip(seq, seq[1:]) if ra[0] != rb[0]]
    print('flips:', flips)
    for x, (tp, np_) in seq:
        print(f'X={x} target_p={tp} pages={np_}')
    return res

if __name__ == '__main__':
    if sys.argv[1] == 'gen':
        gen()
    else:
        measure()
