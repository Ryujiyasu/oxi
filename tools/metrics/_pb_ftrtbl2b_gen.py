# -*- coding: utf-8 -*-
"""Addendum g2: footer_dist(100pt) > bottom_margin(72) so the margin never
binds — distinguishes 'single empty auto para footer = stack 0 (exempt)' from
'stack <= 7.35 (hidden by the margin)'.

  g2_e1n : footer = 1 empty auto para     stack 0 -> cbot 741.9 ; 12.65 -> 729.25
  g2_cP2 : footer = 2 empty auto paras    calibration (expect 25.3 -> 716.6)

K=48 fillers.  keep iff 72 + 48*ADV + X/20 + ADV <= cbot  (thr = 691.80 + X/20)
"""
import os, sys, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs2")
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
           '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr>'
          '</w:style>'
          '</w:styles>')
R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
ADV = 12.6489
SP0 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
EP = f'<w:p><w:pPr>{SP0}</w:pPr></w:p>'
K = 48
FOOTER_TW = 2000   # 100pt > margin 72 -> footer always binds

FOOTERS = {'e1n': EP, 'cP2': EP * 2}
SWEEP = {'e1n': (300, 1100), 'cP2': (300, 700)}


def build(cfg, spacer_tw):
    ps = []
    for i in range(K):
        ps.append(f'<w:p><w:pPr>{SP0}<w:rPr>{R}</w:rPr></w:pPr>'
                  f'<w:r><w:rPr>{R}</w:rPr><w:t>Item {i:02d} alpha beta gamma.</w:t></w:r></w:p>')
    ps.append(f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" '
              f'w:line="{spacer_tw}" w:lineRule="exact"/><w:rPr>{R}</w:rPr></w:pPr></w:p>')
    ps.append(f'<w:p><w:pPr>{SP0}<w:rPr>{R}</w:rPr></w:pPr>'
              f'<w:r><w:rPr>{R}</w:rPr><w:t>TARGETLINE omega.</w:t></w:r></w:p>')
    b = ''.join(ps)
    b += ('<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
          '<w:pgSz w:w="11906" w:h="16838"/>'
          f'<w:pgMar w:top="1440" w:right="1418" w:bottom="1440" w:left="1418" '
          f'w:header="709" w:footer="{FOOTER_TW}" w:gutter="0"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{b}</w:body></w:document>')
    ftr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:ftr {W_NS}>{FOOTERS[cfg]}</w:ftr>')
    return doc, ftr


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    n = 0
    for cfg, (lo, hi) in SWEEP.items():
        for x in range(lo, hi + 1, 50):
            doc, ftr = build(cfg, x)
            with zipfile.ZipFile(os.path.join(OUTDIR, f'g2_{cfg}_{x:04d}.docx'),
                                 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT)
                z.writestr('_rels/.rels', RELS)
                z.writestr('word/_rels/document.xml.rels', DOCRELS)
                z.writestr('word/document.xml', doc)
                z.writestr('word/styles.xml', STYLES)
                z.writestr('word/footer1.xml', ftr)
            n += 1
    print(f'generated {n} -> {os.path.abspath(OUTDIR)}')


def measure():
    import win32com.client
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for fn in sorted(f for f in os.listdir(OUTDIR) if f.endswith('.docx')):
            p = os.path.abspath(os.path.join(OUTDIR, fn))
            d = word.Documents.Open(p, ReadOnly=True)
            try:
                n = d.Paragraphs.Count
                rng = d.Paragraphs(n).Range
                pgno = d.Range(rng.Start, rng.Start).Information(3)
            finally:
                d.Close(False)
            cfg, x = fn[3:-5].rsplit('_', 1)
            res.setdefault(cfg, {})[int(x)] = pgno
            print(f'  {fn}: TARGET page {pgno}', flush=True)
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUTDIR, '_measure.json'), 'w'), indent=1)
    thr0 = 72 + K * ADV + ADV
    print(f'\n=== cbot windows (thr = {thr0:.2f} + X/20; stack = 841.9 - cbot - 100) ===')
    for cfg in FOOTERS:
        if cfg not in res: continue
        d = res[cfg]; xs = sorted(d)
        keep = [x for x in xs if d[x] == 1]; push = [x for x in xs if d[x] > 1]
        if not keep or not push:
            print(f'  {cfg}: NO FLIP ({len(keep)} keep / {len(push)} push)'); continue
        lo, hi = max(keep), min(push)
        c_lo, c_hi = thr0 + lo/20.0, thr0 + hi/20.0
        s_lo, s_hi = 841.9 - c_hi - 100, 841.9 - c_lo - 100
        print(f'  {cfg}: cbot ∈ [{c_lo:.2f}, {c_hi:.2f})  stack ∈ ({s_lo:.2f}, {s_hi:.2f}]'
              f'  ~{(s_lo+s_hi)/2:.2f} ({(s_lo+s_hi)/2/ADV:.2f} lines)')


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == 'gen': gen()
    elif a and a[0] == 'measure': measure()
    else: print(__doc__)
