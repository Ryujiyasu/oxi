"""When does Word DRAW a `default`-footer PAGE field — and does it RESERVE for it?

CONTEXT (2026-07-17, Unit A of the 2基体制). The task premise was that Word
renders NO footer page number on ohnoikuji_03 (a live, default-typed footer with
a PAGE field), while Oxi renders 1..19 => "spurious ink". DIRECT MEASUREMENT
FALSIFIED that premise: Word renders `- 1 -`..`- 19 -` on all 19 pages, at the
SAME position Oxi draws them (x within 0.1pt, y0 780.6 vs 780.7). The "renders
nothing" reading was a FOOTER-BAND ARTIFACT: the digit sits at y0=780.6 on an
H=841.9 page (0.9272*H, only 61.3pt above the bottom edge), so a "footer = last
<=60pt" / ">0.93*H" band MISSES it (a "last 70pt" or ">0.92*H" band catches it).

ohnoikuji_03's footer shape (verified from the bytes):
  - single sectPr, w:type="continuous", footerReference type="default" -> footer1
  - footer1 = one center-jc para, pStyle=a7 (footer, w:semiHidden), rStyle a6
    (page number, w:semiHidden), a WELL-FORMED field:
    "- " + fldChar(begin) + instrText " PAGE " + fldChar(separate) + cached "2"
    + fldChar(end) + " -", TNR, sz(inherited a7)=20 (10pt), black.
  - NO w:vanish, NO evenAndOddHeaders, NO titlePg, not in a textbox/frame.
  => every ECMA condition says Word SHOULD draw it, and it DOES.

This probe establishes the DISCRIMINATOR "does Word DRAW a default-footer PAGE
field", so future footer-suppression rules (the S913 lineage) have a measured
basis, and confirms ohnoikuji's exact shape (continuous + semiHidden + well-
formed) DRAWS. Geometry == ohnoikuji (A4, top=1985 bottom=1701 footer=992 tw) so
the reservation numbers map directly onto the real doc.

VARIANTS (all: A4, default footer, center-jc, TNR sz20, body = NFILL Arial-11
filler lines "Lnnn ..." so page-1 capacity is readable from the PDF):
  base          : plain default section, well-formed "- PAGE -"     -> EXPECT draw
  continuous    : sectPr w:type="continuous" (ohnoikuji shape)      -> EXPECT draw
  semihidden    : footer para pStyle=a7(semiHidden) + rStyle a6     -> EXPECT draw
                  (control: semiHidden is a UI-gallery hint, NOT render suppression)
  vanish        : the PAGE field runs wrapped in <w:vanish/>        -> hyp (a)
  unterminated  : fldChar begin + instrText, NO separate/end        -> hyp (c)
  white         : field run color=FFFFFF                            -> hyp (f)

READOUTS per variant (from the Word PDF, whole-page scan -- NEVER a hard band):
  draw?    : is there an isolated footer digit in the bottom ~90pt?
  ink_y0   : its y0 (should be ~780.6 == ohnoikuji)
  n_cap    : page-1 filler capacity (reservation probe).
             no-reserve cbot = 841.9 - 85.05 = 756.85  (bottom margin binds)
             reserve   cbot = 841.9 - 49.60 - stack; a ~10pt footer stack ~= 11.5
                        -> reserve cbot ~= 780.8 > 756.85 => margin STILL binds
             => for THIS footer geometry reservation does NOT bind (render-only),
                exactly as seen on the real doc (body ends 743.66 < cbot 756.85).
             `tall` control (line=exact 600tw stack=30) makes reserve bind, as the
             validity check that the capacity readout is live.

Usage:
  python _pb_ftrdraw_gen.py gen
  python _pb_ftrdraw_gen.py measure
"""
import os, re, sys, glob, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_ftrdraw")

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

PAGE_H = 841.9
TOP_TW = 1985      # 99.25pt   (== ohnoikuji)
BOTTOM_TW = 1701   # 85.05pt   (== ohnoikuji)  -> no-reserve cbot 756.85
FOOTER_TW = 992    # 49.60pt   (== ohnoikuji)
ADV = 12.6489      # Arial 11 hhea (S805)
NFILL = 80         # spans ~1.3 pages so page-1 capacity is clean

# (case, sect_type, footer_kind, tall)
CASES = [
    ('base',         'nextPage',   'wellformed', False),
    ('continuous',   'continuous', 'wellformed', False),   # ohnoikuji shape
    ('semihidden',   'continuous', 'semihidden', False),   # ohnoikuji styles
    ('vanish',       'nextPage',   'vanish',     False),
    ('unterminated', 'nextPage',   'unterminated', False),
    ('white',        'nextPage',   'white',      False),
    ('tall',         'nextPage',   'wellformed', True),     # reservation validity ctrl
]

R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
SP0 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
# footer field run props: TNR sz20 (10pt) like ohnoikuji
FR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/>'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
           '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
           '</Relationships>')

# styles: Normal(Arial11) + a7 footer(semiHidden, TNR sz20) + a6 pageNumber rStyle(semiHidden)
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="a7"><w:name w:val="footer"/>'
          '<w:semiHidden/><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr></w:style>'
          '<w:style w:type="character" w:styleId="a6"><w:name w:val="page number"/><w:semiHidden/></w:style>'
          '</w:styles>')


def settings():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings {W_NS}></w:settings>')


def fillers(k):
    return ''.join(
        f'<w:p><w:pPr>{SP0}<w:rPr>{R}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{R}</w:rPr><w:t>L{i:03d} alpha beta gamma delta.</w:t></w:r></w:p>'
        for i in range(k))


def sect(stype):
    t = f'<w:type w:val="{stype}"/>' if stype != 'nextPage' else ''
    return ('<w:sectPr>'
            '<w:footerReference w:type="default" r:id="rId2"/>'
            f'{t}<w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="{TOP_TW}" w:right="1701" w:bottom="{BOTTOM_TW}" '
            f'w:left="1701" w:header="851" w:footer="{FOOTER_TW}" w:gutter="0"/>'
            '</w:sectPr>')


def footer_xml(kind, tall):
    """center-jc footer '- PAGE -'. `kind` selects the suppression variant."""
    line = '<w:spacing w:before="0" w:after="0" w:line="600" w:lineRule="exact"/>' if tall else ''
    ppr = (f'<w:pPr>{line}'
           + ('<w:pStyle w:val="a7"/>' if kind == 'semihidden' else '')
           + '<w:jc w:val="center"/></w:pPr>')

    field_rpr = FR
    if kind == 'white':
        field_rpr = FR + '<w:color w:val="FFFFFF"/>'
    if kind == 'semihidden':
        field_rpr = FR + '<w:rStyle w:val="a6"/>'
    if kind == 'vanish':
        field_rpr = FR + '<w:vanish/>'

    def run(inner):
        return f'<w:r><w:rPr>{field_rpr}</w:rPr>{inner}</w:r>'

    if kind == 'unterminated':
        # begin + instrText, NO separate, NO end -> a broken/unterminated field
        field = (run('<w:fldChar w:fldCharType="begin"/>')
                 + run('<w:instrText xml:space="preserve"> PAGE </w:instrText>'))
    else:
        field = (run('<w:fldChar w:fldCharType="begin"/>')
                 + run('<w:instrText xml:space="preserve"> PAGE </w:instrText>')
                 + run('<w:fldChar w:fldCharType="separate"/>')
                 + run('<w:t>2</w:t>')                       # cached result
                 + run('<w:fldChar w:fldCharType="end"/>'))

    dash1 = f'<w:r><w:rPr>{field_rpr}</w:rPr><w:t xml:space="preserve">- </w:t></w:r>'
    dash2 = f'<w:r><w:rPr>{field_rpr}</w:rPr><w:t xml:space="preserve"> -</w:t></w:r>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:ftr {W_NS}><w:p>{ppr}{dash1}{field}{dash2}</w:p></w:ftr>')


def build(case, stype, kind, tall):
    body = fillers(NFILL) + sect(stype)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')
    return doc, footer_xml(kind, tall)


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for case, stype, kind, tall in CASES:
        doc, ftr = build(case, stype, kind, tall)
        p = os.path.join(OUTDIR, f'fd_{case}.docx')
        with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', doc)
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/settings.xml', settings())
            z.writestr('word/footer1.xml', ftr)
    print(f'generated {len(CASES)} -> {os.path.abspath(OUTDIR)}')


def readout(pdf):
    """page-1: footer digit (draw?), its y0, and filler capacity n."""
    import fitz
    d = fitz.open(pdf)
    pg = d[0]
    H = pg.rect.height
    ncap = 0
    digit = None
    dy = None
    for blk in pg.get_text('dict')['blocks']:
        if blk.get('type') != 0:
            continue
        for ln in blk['lines']:
            spans = ln['spans']
            t = ''.join(s['text'] for s in spans)
            if re.match(r'^L\d{3}\b', t):
                ncap += 1
            # footer line: any line in the bottom ~90pt whose text contains a
            # digit (Word may merge "- 1 -" into one span OR split it).
            y0 = min(s['bbox'][1] for s in spans)
            if y0 > H - 90:
                m = re.search(r'\d+', t)
                if m:
                    digit = m.group(0)
                    dy = round(y0, 2)
    npages = d.page_count
    d.close()
    return {'draw': digit is not None, 'digit': digit, 'ink_y0': dy,
            'n_cap': ncap, 'pages': npages}


def measure():
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for case, *_ in CASES:
            f = os.path.join(OUTDIR, f'fd_{case}.docx')
            pdf = f[:-5] + '.pdf'
            if not os.path.exists(pdf):
                dd = word.Documents.Open(os.path.abspath(f), ReadOnly=True)
                dd.ExportAsFixedFormat(os.path.abspath(pdf), 17)
                dd.Close(False)
            res[case] = readout(pdf)
            print('  measured', case, res[case], flush=True)
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUTDIR, '_measure.json'), 'w'), indent=1)
    report(res)


def report(res=None):
    if res is None:
        res = json.load(open(os.path.join(OUTDIR, '_measure.json')))
    no_reserve_cbot = PAGE_H - BOTTOM_TW / 20.0
    n_margin = int((no_reserve_cbot - TOP_TW / 20.0) // ADV)
    print('\n=== geometry ===')
    print(f'  H={PAGE_H}  top={TOP_TW/20:.2f}  bottom={BOTTOM_TW/20:.2f}  footer_dist={FOOTER_TW/20:.2f}')
    print(f'  no-reserve cbot = {no_reserve_cbot:.2f}  -> page-1 filler capacity n = {n_margin}')
    print('\n=== results ===')
    print(f"  {'case':13} {'draw':5} {'digit':6} {'ink_y0':8} {'n_cap':6} {'pages':6}")
    for case, *_ in CASES:
        r = res.get(case, {})
        print(f"  {case:13} {str(r.get('draw')):5} {str(r.get('digit')):6} "
              f"{str(r.get('ink_y0')):8} {str(r.get('n_cap')):6} {str(r.get('pages')):6}")


if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if cmd == 'gen':
        gen()
    elif cmd == 'measure':
        measure()
    elif cmd == 'report':
        report()
    else:
        print(__doc__)
