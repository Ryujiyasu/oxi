# -*- coding: utf-8 -*-
"""Footer TABLE stack: the BORDER / cellMar / trHeight terms (+ the cP1
blank-footer outlier discriminator).  Extends _pb_ftrtbl_gen (which proved
cT1==cP2 / cT2==cP3: an unbordered footer table costs exactly its rows).

Open questions this pins:
  1. forms__000ee7c0's Word footer row renders 12.72 for an ~9.8pt cell line
     ("double-border box") — do the row's TOP/BOTTOM borders add to the
     footer keep-out stack, and at what width (single sz4/sz12, double sz4)?
  2. does tblCellMar top/bottom add? (TableNormal default is 0/0, but real
     docs declare it)
  3. does trHeight (atLeast) drive the reserved row height?
  4. cP1 outlier: a footer of ONE empty auto-line para reserved NOTHING
     (cbot = margin) while 2 empty paras / 1 exact-line para push fully.
     Discriminators tested: ink (e1i), mark size (e1s28), replication (e1n).

Geometry / method = _pb_ftrtbl_gen verbatim: A4, bottom=1440tw, footer=1293tw,
K=50 Arial-11 filler lines + ONE line=X exact spacer + TARGET line; the
TARGET p1->p2 flip pins cbot:
    keep iff 72 + K*ADV + X/20 + ADV <= cbot
    stack = 841.9 - cbot - 64.65    (validated: reserved = max(72, 64.65+stack))

Usage:
  python _pb_ftrtbl2_gen.py gen
  python _pb_ftrtbl2_gen.py measure
  python _pb_ftrtbl2_gen.py report
"""
import os, sys, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs")

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
ADV = 12.6489      # Arial 11 hhea (S805)
SP0 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
EP = f'<w:p><w:pPr>{SP0}</w:pPr></w:p>'
K = 50


def tbl(borders=None, cellmar=None, trh=None):
    """1-row x 3-cell table, each cell = one empty spacing-0 para.
    borders: None | (val, sz) applied to top/left/bottom/right (+insideV)
    cellmar: None | (top_tw, bottom_tw)
    trh:     None | (val_tw, hrule)
    """
    bx = ''
    if borders:
        v, sz = borders
        edge = lambda e: f'<w:{e} w:val="{v}" w:sz="{sz}" w:space="0" w:color="auto"/>'
        bx = ('<w:tblBorders>' + edge('top') + edge('left') + edge('bottom')
              + edge('right')
              + f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>'
              + '</w:tblBorders>')
    cm = ''
    if cellmar:
        t, b = cellmar
        cm = (f'<w:tblCellMar><w:top w:w="{t}" w:type="dxa"/>'
              f'<w:bottom w:w="{b}" w:type="dxa"/></w:tblCellMar>')
    trpr = ''
    if trh:
        val, rule = trh
        trpr = f'<w:trPr><w:trHeight w:val="{val}" w:hRule="{rule}"/></w:trPr>'
    cell = f'<w:tc><w:tcPr><w:tcW w:w="3020" w:type="dxa"/></w:tcPr>{EP}</w:tc>'
    tr = f'<w:tr>{trpr}{cell*3}</w:tr>'
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
            + bx + '<w:tblLayout w:type="fixed"/>' + cm + '</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3020"/><w:gridCol w:w="3020"/>'
            '<w:gridCol w:w="3020"/></w:tblGrid>'
            + tr + '</w:tbl>')


EP_INK = f'<w:p><w:pPr>{SP0}<w:rPr>{R}</w:rPr></w:pPr><w:r><w:rPr>{R}</w:rPr><w:t>F</w:t></w:r></w:p>'
EP_S28 = (f'<w:p><w:pPr>{SP0}<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
          '<w:sz w:val="28"/></w:rPr></w:pPr></w:p>')

FOOTERS = {
    # recalibration anchor (prior run: flip (680,700] -> stack ~25.3)
    'cP2r':  EP * 2,
    # border terms
    'bS4':   tbl(borders=('single', 4)) + EP,
    'bS12':  tbl(borders=('single', 12)) + EP,
    'bD4':   tbl(borders=('double', 4)) + EP,      # the forms shape
    # cellMar term
    'cM105': tbl(cellmar=(105, 105)) + EP,
    # additivity
    'bD4M':  tbl(borders=('double', 4), cellmar=(105, 105)) + EP,
    # trHeight term
    'tRH':   tbl(trh=(500, 'atLeast')) + EP,
    # cP1 outlier discriminators (single footer para)
    'e1n':   EP,          # replication of cP1
    'e1i':   EP_INK,      # + ink
    'e1s':   EP_S28,      # empty, mark sz=28 (14pt)
}

SWEEP = {
    'cP2r':  (620, 780),
    'bS4':   (560, 760),
    'bS12':  (500, 720),
    'bD4':   (500, 760),
    'cM105': (380, 600),
    'bD4M':  (320, 560),
    'tRH':   (340, 560),
    'e1n':   (880, 1120),
    'e1i':   (880, 1120),
    'e1s':   (800, 1120),
}


def body(k, spacer_tw):
    ps = []
    for i in range(k):
        ps.append(f'<w:p><w:pPr>{SP0}<w:rPr>{R}</w:rPr></w:pPr>'
                  f'<w:r><w:rPr>{R}</w:rPr><w:t>Item {i:02d} alpha beta gamma.</w:t></w:r></w:p>')
    ps.append(f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" '
              f'w:line="{spacer_tw}" w:lineRule="exact"/><w:rPr>{R}</w:rPr></w:pPr></w:p>')
    ps.append(f'<w:p><w:pPr>{SP0}<w:rPr>{R}</w:rPr></w:pPr>'
              f'<w:r><w:rPr>{R}</w:rPr><w:t>TARGETLINE omega.</w:t></w:r></w:p>')
    return ''.join(ps)


def build(cfg, spacer_tw):
    b = body(K, spacer_tw)
    b += ('<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
          '<w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="1440" w:right="1418" w:bottom="1440" w:left="1418" '
          'w:header="709" w:footer="1293" w:gutter="0"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{b}</w:body></w:document>')
    ftr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:ftr {W_NS}>{FOOTERS[cfg]}</w:ftr>')
    return doc, ftr


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    n = 0
    for cfg, (lo, hi) in SWEEP.items():
        for x in range(lo, hi + 1, 20):
            doc, ftr = build(cfg, x)
            p = os.path.join(OUTDIR, f'f2_{cfg}_{x:04d}.docx')
            with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT)
                z.writestr('_rels/.rels', RELS)
                z.writestr('word/_rels/document.xml.rels', DOCRELS)
                z.writestr('word/document.xml', doc)
                z.writestr('word/styles.xml', STYLES)
                z.writestr('word/footer1.xml', ftr)
            n += 1
    print(f'generated {n} -> {os.path.abspath(OUTDIR)}')


def cbot_of(x):
    return 72 + K * ADV + x / 20.0 + ADV


def measure():
    import win32com.client
    # DispatchEx: PRIVATE Word instance (do not attach to any instance the
    # parallel agent may be using)
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    mp = os.path.join(OUTDIR, '_measure.json')
    if os.path.exists(mp):
        res = json.load(open(mp))
    try:
        files = sorted(f for f in os.listdir(OUTDIR) if f.endswith('.docx'))
        for fn in files:
            cfg, x = fn[3:-5].rsplit('_', 1)
            if res.get(cfg, {}).get(x) is not None:
                continue
            p = os.path.abspath(os.path.join(OUTDIR, fn))
            d = word.Documents.Open(p, ReadOnly=True)
            try:
                n = d.Paragraphs.Count
                rng = d.Paragraphs(n).Range
                r0 = d.Range(rng.Start, rng.Start)
                pgno = r0.Information(3)
            finally:
                d.Close(False)
            res.setdefault(cfg, {})[x] = pgno
            print(f'  {fn}: TARGET page {pgno}', flush=True)
            json.dump(res, open(mp, 'w'), indent=1)
    finally:
        word.Quit()
    report()


def report():
    res = json.load(open(os.path.join(OUTDIR, '_measure.json')))
    print('\n=== stack per config (reserved = max(72, 64.65 + stack)) ===')
    print(f'    para line ADV = {ADV}')
    for cfg in FOOTERS:
        if cfg not in res:
            continue
        d = {int(k): v for k, v in res[cfg].items()}
        xs = sorted(d)
        keep = [x for x in xs if d[x] == 1]
        push = [x for x in xs if d[x] > 1]
        if not keep or not push:
            print(f'  {cfg}: NO FLIP in range ({len(keep)} keep / {len(push)} push)')
            continue
        lo, hi = max(keep), min(push)
        c_lo, c_hi = cbot_of(lo), cbot_of(hi)
        s_lo, s_hi = 841.9 - c_hi - 64.65, 841.9 - c_lo - 64.65
        mid = (s_lo + s_hi) / 2
        print(f'  {cfg:6s}: cbot ∈ [{c_lo:.2f}, {c_hi:.2f})  stack ∈ ({s_lo:.2f}, {s_hi:.2f}] '
              f' ~{mid:.2f}  ({mid/ADV:.2f} lines)')


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == 'gen':
        gen()
    elif a and a[0] == 'measure':
        measure()
    elif a and a[0] == 'report':
        report()
    else:
        print(__doc__)
