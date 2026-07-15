"""When does a FOOTER push the body up? (the S868 blocker)

_pb_ftrtbl_gen proved a footer TABLE costs exactly its rows (cT1 ≡ cP2,
cT2 ≡ cP3). But the same sweep exposed a bigger unknown: a footer does NOT
always push the body.

    cP1 (1 line, stack 12.65): footer_dist 64.65 + 12.65 = 77.3 > margin 72
        -> the footer's top (764.6) is 5.3pt ABOVE the margin line (769.9),
           yet Word kept cbot at the MARGIN (769.9) = intrusion IGNORED.
    cP2 (2 lines, stack 25.30): intrusion 17.95 -> Word pushed, cbot 751.6
           = the footer's top EXACTLY = intrusion RESPECTED.
    cP3 (3 lines): pushed likewise.

And the real doc forms__000ee7c0 (footer_dist 14.4, margin 21.6, a table+para
footer that RENDERS both rows -- marker-pinned) behaves like cP1: Word keeps
the body down at ~the margin even though the footer intrudes ~16pt, which is
why S868's (correct) table height OVER-reserves there and cost it a PASS.

So the reservation model `reserved = max(bottom_margin, footer_dist + stack)`
is NOT what Word does at small stacks. This probe pins the transition.

METHOD (cheap: 1 render per config): body = N plain filler lines, footer = ONE
paragraph with line=X lineRule=exact, so the footer stack is EXACTLY X/20 pt
and sweeps continuously. Read the page-1 filler count n:
    cbot ∈ [72 + n*ADV, 72 + (n+1)*ADV)
    no-push model  -> cbot = 769.9            (n = 55)
    full-push model-> cbot = 777.25 - X/20    (n shrinks as X grows)
The X at which n starts shrinking IS the threshold.

Usage:
  python _pb_ftrpush_gen.py gen [lo hi step]
  python _pb_ftrpush_gen.py measure
"""
import os, sys, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_ftrpush")

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
          '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr>'
          '</w:style></w:styles>')

R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
SP0 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
ADV = 12.6489
NFILL = 62
FOOTER_TW = 1293   # 64.65pt (administrative__0006985e)
BOTTOM_TW = 1440   # 72pt
PAGE_H = 841.9


def build(x_tw, ink):
    """footer = ONE para, line=x_tw exact. ink=True -> it paints a glyph."""
    ps = ''.join(
        f'<w:p><w:pPr>{SP0}<w:rPr>{R}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{R}</w:rPr><w:t>L{i:02d} alpha beta gamma.</w:t></w:r></w:p>'
        for i in range(NFILL))
    body = ps + (f'<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
                 f'<w:pgSz w:w="11906" w:h="16838"/>'
                 f'<w:pgMar w:top="1440" w:right="1418" w:bottom="{BOTTOM_TW}" '
                 f'w:left="1418" w:header="709" w:footer="{FOOTER_TW}" '
                 f'w:gutter="0"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')
    run = f'<w:r><w:rPr>{R}</w:rPr><w:t>F</w:t></w:r>' if ink else ''
    ftr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:ftr {W_NS}><w:p><w:pPr>'
           f'<w:spacing w:before="0" w:after="0" w:line="{x_tw}" w:lineRule="exact"/>'
           f'<w:rPr>{R}</w:rPr></w:pPr>{run}</w:p></w:ftr>')
    return doc, ftr


def name(x, ink):
    return f"fp_{'ink' if ink else 'nil'}_{x:04d}"


def gen(cases):
    os.makedirs(OUTDIR, exist_ok=True)
    for x, ink in cases:
        doc, ftr = build(x, ink)
        with zipfile.ZipFile(os.path.join(OUTDIR, name(x, ink) + '.docx'),
                             'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', doc)
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/footer1.xml', ftr)
    print(f'generated {len(cases)} -> {os.path.abspath(OUTDIR)}')


def measure():
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for fn in sorted(f for f in os.listdir(OUTDIR) if f.endswith('.docx')):
            p = os.path.abspath(os.path.join(OUTDIR, fn))
            d = word.Documents.Open(p, ReadOnly=True)
            try:
                n1 = 0
                for i in range(1, d.Paragraphs.Count + 1):
                    rng = d.Paragraphs(i).Range
                    if d.Range(rng.Start, rng.Start).Information(3) == 1:
                        n1 += 1
                    else:
                        break
            finally:
                d.Close(False)
            res[fn[:-5]] = n1
            print(f'  {fn[:-5]}: {n1} lines on page 1', flush=True)
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUTDIR, '_measure.json'), 'w'), indent=1)
    print('\n=== footer stack X -> page-1 capacity -> implied cbot ===')
    print(f'  margin cbot = {PAGE_H - BOTTOM_TW/20:.2f} (no-push model)')
    for ink in (False, True):
        print(f'  --- footer {"WITH ink" if ink else "EMPTY (no ink)"}')
        for k in sorted(res):
            if k.startswith('fp_ink_') != ink:
                continue
            x = int(k.rsplit('_', 1)[1])
            n = res[k]
            lo = 72 + n * ADV
            full_push = PAGE_H - FOOTER_TW/20 - x/20.0
            print(f'    stack={x/20:5.2f}pt  n={n:2d}  cbot∈[{lo:.1f},{lo+ADV:.1f})'
                  f'   full-push predicts {full_push:.1f}'
                  f'   {"PUSHED" if lo < PAGE_H - BOTTOM_TW/20 - ADV else "no push"}')


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == 'gen':
        lo, hi, st = (int(a[1]), int(a[2]), int(a[3])) if len(a) > 3 else (100, 760, 40)
        gen([(x, ink) for x in range(lo, hi + 1, st) for ink in (False, True)])
    elif a and a[0] == 'measure':
        measure()
    else:
        print(__doc__)
