# -*- coding: utf-8 -*-
"""ctx-spacing rule, round 2 — decide between:
  OWNER model : gap = max(A.after, B.before); removed ENTIRELY iff
                (A.ctx && A.after >= B.before) || (B.ctx && B.before >= A.after)
  NONZERO model: removed entirely iff (A.ctx && A.after > 0) || (B.ctx && B.before > 0)
Both fit educational (A ctx 180/180 -> gap 0) and c1..c6.

  d1: A ctx after=90 , B before=360          OWNER: keep 18   NONZERO: 0
  d2: A ctx after=360, B before=90           both: 0
  d3: A ctx after=360, B before=180          total: 0   partial-recompute: 9
  d4: A ctx after=180, B before=180          educational-minimal: expect 0
  d5: A ctx after=180, B before=180, line=360 both paras (edu closer)  expect 0 (+taller line)
  d6: A plain after=0, B ctx before=360      textbook lower-own: 0
  d7: A plain after=360, B ctx before=0      B removes own(0) -> A.after survives? keep 18
  d8: A ctx after=360, B pStyle=Alt before=0 diff style -> ctx disabled -> keep 18
gap printed = Info6(B) - Info6(A) = A's line box + applied inter-spacing.
line=240 -> 12.65; line=360 -> 18.97 (Arial 11 hhea x1.5).
"""
import os, sys, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs4")
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/>'
          '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr>'
          '</w:style>'
          '<w:style w:type="paragraph" w:styleId="Alt"><w:name w:val="Alt"/>'
          '<w:basedOn w:val="Normal"/>'
          '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '</w:style>'
          '</w:styles>')
R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
SP0 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'


def sp(before=0, after=0, line=240):
    return f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'


CTX = '<w:contextualSpacing/>'
CONFIGS = {
    'd1': (sp(after=90) + CTX,            sp(before=360),        None),
    'd2': (sp(after=360) + CTX,           sp(before=90),         None),
    'd3': (sp(after=360) + CTX,           sp(before=180),        None),
    'd4': (sp(after=180) + CTX,           sp(before=180),        None),
    'd5': (sp(after=180, line=360) + CTX, sp(before=180, line=360), None),
    'd6': (SP0,                           sp(before=360) + CTX,  None),
    'd7': (sp(after=360),                 sp(before=0) + CTX,    None),
    'd8': (sp(after=360) + CTX,           '<w:pStyle w:val="Alt"/>' + sp(before=0), None),
}


def para(text, ppr):
    return (f'<w:p><w:pPr>{ppr}</w:pPr>'
            f'<w:r><w:rPr>{R}</w:rPr><w:t>{text}</w:t></w:r></w:p>')


def build(cfg):
    pa, pb, _ = CONFIGS[cfg]
    body = (para('Filler zero line.', SP0)
            + para('AAA upper paragraph.', pa)
            + para('BBB lower paragraph.', pb)
            + para('CCC tail.', SP0)
            + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="1440" w:right="1418" w:bottom="1440" w:left="1418" '
              'w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for cfg in CONFIGS:
        with zipfile.ZipFile(os.path.join(OUTDIR, f'cs_{cfg}.docx'), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', build(cfg))
            z.writestr('word/styles.xml', STYLES)
    print(f'generated {len(CONFIGS)}')


def measure():
    import win32com.client
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    out = {}
    try:
        for fn in sorted(f for f in os.listdir(OUTDIR) if f.endswith('.docx')):
            d = word.Documents.Open(os.path.abspath(os.path.join(OUTDIR, fn)), ReadOnly=True)
            try:
                ys = []
                for i in range(1, d.Paragraphs.Count + 1):
                    rng = d.Paragraphs(i).Range
                    ys.append(d.Range(rng.Start, rng.Start).Information(6))
                gab = ys[2] - ys[1]
                gbc = ys[3] - ys[2]
                out[fn[3:-5]] = {'gapAB': round(gab, 2), 'gapBC': round(gbc, 2)}
                print(f'  {fn[3:-5]}: gap A->B = {gab:6.2f}   gap B->C = {gbc:6.2f}')
            finally:
                d.Close(False)
    finally:
        word.Quit()
    json.dump(out, open(os.path.join(OUTDIR, '_measure.json'), 'w'), indent=1)


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == 'gen': gen()
    elif a and a[0] == 'measure': measure()
    else: print(__doc__)
