"""Footer TABLE term derivation (the S868 residual).

S868 made a footer <w:tbl> contribute its row heights, but administrative__
0006985e still needs ~13pt MORE: its footer RENDERS 2 lines (a 1-row 3-cell
table at ink 752.34 + a trailing para at 765.06, pitch 12.72 — pinned by
injecting markers) yet Word's body cbot is bounded to [734.5, 743.5), i.e. an
implied stack of ~34-43 ~= THREE lines. So Word appears to add a per-TABLE
term to the footer keep-out that the (validated) S806 paragraph stack lacks.

DECISIVE COMPARISON: footers that RENDER THE SAME number of lines, one built
from paragraphs and one containing a table:
    cP1 = 1 empty para                       (1 line)
    cP2 = 2 empty paras                      (2 lines)
    cP3 = 3 empty paras                      (3 lines)   <- per-line calibration
    cT1 = table(1 row x 3 cells) + 1 para    (2 lines)   <- compare to cP2
    cT2 = table(2 rows x 3 cells) + 1 para   (3 lines)   <- compare to cP3
If cbot(cT1) == cbot(cP2) the table costs exactly its rows (S868 complete, and
administrative's residual is something else). If cbot(cT1) == cbot(cP3) Word
adds ~one extra line per table.

Geometry = administrative__0006985e's (A4, bottom=1440tw=72, footer=1293tw=
64.65) so the footer BINDS (footer_dist + stack > bottom_margin) — the
_pb_fstack default (footer=709) is margin-bound and hides the footer entirely.

Method (the _pb_fstack exact-spacer sweep): K filler lines + ONE empty spacer
para with line=X lineRule=exact + a TARGET line; sweep X; the p1->p2 flip of
TARGET pins cbot to the step:
    keep iff 72 + K*ADV + X/20 + ADV <= cbot      (S779/S827 full hhea line)
    cbot ∈ [72 + K*ADV + Xlast/20 + ADV,  72 + K*ADV + Xfirst/20 + ADV)

Usage:
  python _pb_ftrtbl_gen.py gen [coarse|fine:LO:HI:STEP]
  python _pb_ftrtbl_gen.py measure
"""
import os, sys, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_ftrtbl")

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

# administrative__0006985e's shape: Normal has after=0 line=240 auto (so the
# footer paras are pure line-height, no spacing to confound the stack).
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
EP = f'<w:p><w:pPr>{SP0}</w:pPr></w:p>'      # empty spacing-0 para


def tbl(rows):
    """1..n-row x 3-cell table, each cell = one empty spacing-0 para."""
    cell = f'<w:tc><w:tcPr><w:tcW w:w="3020" w:type="dxa"/></w:tcPr>{EP}</w:tc>'
    tr = f'<w:tr>{cell*3}</w:tr>'
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3020"/><w:gridCol w:w="3020"/>'
            '<w:gridCol w:w="3020"/></w:tblGrid>'
            + tr * rows + '</w:tbl>')


FOOTERS = {
    'cP1': EP,
    'cP2': EP * 2,
    'cP3': EP * 3,
    'cT1': tbl(1) + EP,
    'cT2': tbl(2) + EP,
}
# rendered line counts (for reading the result)
LINES = {'cP1': 1, 'cP2': 2, 'cP3': 3, 'cT1': 2, 'cT2': 3}

K = 50  # filler lines -> cbot = 717.09 + X/20 covers 727..772 for X 200..1100


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


def gen(cases):
    os.makedirs(OUTDIR, exist_ok=True)
    for cfg, x in cases:
        doc, ftr = build(cfg, x)
        p = os.path.join(OUTDIR, f'ft_{cfg}_{x:04d}.docx')
        with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', doc)
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/footer1.xml', ftr)
    print(f'generated {len(cases)} -> {os.path.abspath(OUTDIR)}')


def cbot_of(x):
    return 72 + K * ADV + x / 20.0 + ADV


def measure(pattern=''):
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        files = sorted(f for f in os.listdir(OUTDIR)
                       if f.endswith('.docx') and pattern in f)
        for fn in files:
            p = os.path.abspath(os.path.join(OUTDIR, fn))
            d = word.Documents.Open(p, ReadOnly=True)
            try:
                # page of the TARGET paragraph (the last one)
                n = d.Paragraphs.Count
                rng = d.Paragraphs(n).Range
                r0 = d.Range(rng.Start, rng.Start)
                pgno = r0.Information(3)
            finally:
                d.Close(False)
            cfg, x = fn[3:-5].rsplit('_', 1)
            res.setdefault(cfg, {})[int(x)] = pgno
            print(f'  {fn}: TARGET on page {pgno}', flush=True)
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUTDIR, '_measure.json'), 'w'), indent=1)
    print('\n=== cbot per config (footer stack = 841.9 - cbot - 64.65) ===')
    base = None
    for cfg in ('cP1', 'cP2', 'cP3', 'cT1', 'cT2'):
        if cfg not in res:
            continue
        xs = sorted(res[cfg])
        keep = [x for x in xs if res[cfg][x] == 1]
        push = [x for x in xs if res[cfg][x] > 1]
        if not keep or not push:
            print(f'  {cfg}: no flip in range (keep={len(keep)} push={len(push)})')
            continue
        lo, hi = max(keep), min(push)
        c_lo, c_hi = cbot_of(lo), cbot_of(hi)
        stack_lo, stack_hi = 841.9 - c_hi - 64.65, 841.9 - c_lo - 64.65
        print(f'  {cfg} ({LINES[cfg]} rendered lines): cbot ∈ [{c_lo:.2f}, {c_hi:.2f}) '
              f'-> stack ∈ ({stack_lo:.2f}, {stack_hi:.2f}]  ~{(stack_lo+stack_hi)/2:.2f}'
              f'  = {((stack_lo+stack_hi)/2)/ADV:.2f} lines')


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == 'gen':
        mode = a[1] if len(a) > 1 else 'coarse'
        if mode == 'coarse':
            cases = [(c, x) for c in FOOTERS for x in range(200, 1101, 20)]
        else:
            _, lo, hi, st = mode.split(':')
            cases = [(c, x) for c in FOOTERS
                     for x in range(int(lo), int(hi) + 1, int(st))]
        gen(cases)
    elif a and a[0] == 'measure':
        measure(a[1] if len(a) > 1 else '')
    else:
        print(__doc__)
