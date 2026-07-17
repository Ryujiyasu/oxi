"""Does Word RESERVE a body band for an INERT `type="even"` footer? (S912 blocker)

ECMA-376 §17.10.2: a `<w:footerReference w:type="even">` is only active when
settings.xml carries `<w:evenAndOddHeaders/>`. Several corpus docs (the whole
tokumei_08_01 family + order_06 + kyodoken07) reference ONLY an even footer and
have NO such flag. Word render-truth confirms it DRAWS NOTHING (word_png of
6514f214e482_tokumei_08_01-2 has zero footer ink on all 7 pages).

Oxi renders it anyway: parser/ooxml.rs `pick_fallback` deliberately keeps "even"
as a last-resort fallback, its own comment saying legacy docs "relied on Oxi's
prior all-refs fallback for body-area RESERVATION — keep that fallback to avoid
pagination regressions". So the open question is precisely:

    does Word reserve a body-area band for a footer it does not draw?

METHOD (clone of _pb_ftrpush_gen.py, 1 render per config): body = NFILL plain
filler lines "Lnn ..." (Arial 11, spacing 0, single -> ADV = 12.6489, S805);
footer = ONE paragraph with line=X lineRule=exact carrying an ink glyph, so the
footer stack is EXACTLY X/20 pt and sweeps continuously. Read the per-page
filler capacity out of the PDF (fitz) -- NEVER COM Information(6), which is
0.75pt-quantized (documented trap).

    cbot ∈ [72 + n*ADV, 72 + (n+1)*ADV)
    no reservation   -> cbot = 769.90 (bottom margin)     -> n = 55
    full reservation -> cbot = 841.9 - 64.65 - X/20       -> n shrinks with X

The footer para carries INK on purpose: an ink-free footer para hits the
documented blank-footer exemption (_pb_fstack_gen case fE10 measured stack = 0)
and would make every case degenerate. The real docs' footers carry a PAGE field
(= ink), so ink-bearing is the right analogue.

CASES (page geometry A4, top=1440 bottom=1440(72pt) footer=1293(64.65pt), so the
footer binds as soon as stack > 7.35pt):
  ctrl_none    : no footerReference at all              (capacity baseline)
  even_only    : ONLY type="even", NO evenAndOddHeaders  (the doc under test)
  even_flagged : same + <w:evenAndOddHeaders/>           (VALIDITY control: if
                 ink appears on the EVEN page, the even ref is well-formed, so
                 even_only drawing nothing is Word's rule, not a broken probe)
  default_only : ONLY type="default"                     (reserve+draw control)

DECISIVE READOUT: is even_only's page-1 capacity == ctrl_none's (=> NO
reservation) or == default_only's (=> reservation without draw)?

SECOND QUESTION (the ukframework shape). The corpus scan found that the trigger
also fires on a MULTI-section doc: uk_framework sec2/sec3 declare ONLY an even
footer after a sec1 that declared a default. Oxi's inheritance (ooxml.rs:84-93)
is ALL-OR-NOTHING per section -- `if section.footer_refs.is_empty() { prev }` --
so a section declaring ANY ref inherits NOTHING. ECMA-376 §17.10.1 inherits PER
TYPE. Case `inherit_even` decides it:
  sec1 = A-lines + footerReference default -> footer1 ("FTRONE")
  sec2 = B-lines + footerReference even    -> footer2 ("FTRTWO"), no flag
On sec2's first page:
  FTRONE ink + pushed capacity -> Word INHERITS the default per type
  FTRTWO ink                   -> Word honours the inert even ref (falsifies all)
  no ink  + capacity 55        -> all-or-nothing AND even is inert

Usage:
  python _pb_evenftr_gen.py gen [x_tw ...]      (default 400 1200 2400)
  python _pb_evenftr_gen.py measure
"""
import os, re, sys, glob, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_evenftr")

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

PAGE_H = 841.9
BOTTOM_TW = 1440   # 72.00pt
FOOTER_TW = 1293   # 64.65pt  (== _pb_ftrpush_gen; binds when stack > 7.35)
ADV = 12.6489      # Arial 11 hhea (S805)
NFILL = 130        # spans ~3 pages so page-2 capacity is also measurable

CASES = ('ctrl_none', 'even_only', 'even_flagged', 'default_only', 'inherit_even')
FTR_TYPE = {'ctrl_none': None, 'even_only': 'even',
            'even_flagged': 'even', 'default_only': 'default', 'inherit_even': 'even'}
FLAG = {'ctrl_none': False, 'even_only': False,
        'even_flagged': True, 'default_only': False, 'inherit_even': False}
K1 = 60            # inherit_even: sec1 A-lines (spans 2 pages)
K2 = 70            # inherit_even: sec2 B-lines

R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
SP0 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'

CT_HEAD = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
           '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
           '<Default Extension="xml" ContentType="application/xml"/>'
           '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
           '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
           '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>')
CT_FTR = ('<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>')
CT_FTR2 = ('<Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>')
CT_TAIL = '</Types>'

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOCRELS_HEAD = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>')
DOCRELS_FTR = ('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>')
DOCRELS_FTR2 = ('<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>')
DOCRELS_TAIL = '</Relationships>'

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


def settings(flag):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings {W_NS}>'
            + ('<w:evenAndOddHeaders/>' if flag else '')
            + '</w:settings>')


def fillers(prefix, k):
    return ''.join(
        f'<w:p><w:pPr>{SP0}<w:rPr>{R}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{R}</w:rPr><w:t>{prefix}{i:03d} alpha beta gamma delta.</w:t></w:r></w:p>'
        for i in range(k))


def sect(ref, final=True):
    s = (f'<w:sectPr>{ref}'
         f'<w:pgSz w:w="11906" w:h="16838"/>'
         f'<w:pgMar w:top="1440" w:right="1418" w:bottom="{BOTTOM_TW}" '
         f'w:left="1418" w:header="709" w:footer="{FOOTER_TW}" '
         f'w:gutter="0"/></w:sectPr>')
    # a non-final sectPr lives inside the last paragraph of its section
    return s if final else f'<w:p><w:pPr>{s}</w:pPr></w:p>'


def footer_xml(x_tw, ink):
    """ONE para, line=x exact -> stack == x/20 pt exactly. INK on purpose (an
    empty footer para hits the blank-footer exemption, fE10 stack = 0)."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:ftr {W_NS}><w:p><w:pPr>'
            f'<w:spacing w:before="0" w:after="0" w:line="{x_tw}" w:lineRule="exact"/>'
            f'<w:rPr>{R}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R}</w:rPr><w:t>{ink}</w:t></w:r></w:p></w:ftr>')


def build(case, x_tw):
    """-> (document.xml, {part: xml}, n_footers)"""
    ftype = FTR_TYPE[case]
    if case == 'inherit_even':
        # sec1: default footer1 (FTRONE) | sec2: even footer2 (FTRTWO), no flag
        body = (fillers('A', K1)
                + sect('<w:footerReference w:type="default" r:id="rId2"/>', final=False)
                + fillers('B', K2)
                + sect('<w:footerReference w:type="even" r:id="rId4"/>', final=True))
        parts = {'word/footer1.xml': footer_xml(x_tw, 'FTRONE'),
                 'word/footer2.xml': footer_xml(x_tw, 'FTRTWO')}
    else:
        ref = f'<w:footerReference w:type="{ftype}" r:id="rId2"/>' if ftype else ''
        body = fillers('L', NFILL) + sect(ref, final=True)
        parts = {'word/footer1.xml': footer_xml(x_tw, 'FTRINK')} if ftype else {}
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')
    return doc, parts


def gen(xs):
    os.makedirs(OUTDIR, exist_ok=True)
    n = 0
    for case in CASES:
        for x in xs:
            doc, parts = build(case, x)
            ct = CT_HEAD + ('word/footer1.xml' in parts and CT_FTR or '') \
                 + ('word/footer2.xml' in parts and CT_FTR2 or '') + CT_TAIL
            dr = DOCRELS_HEAD + ('word/footer1.xml' in parts and DOCRELS_FTR or '') \
                 + ('word/footer2.xml' in parts and DOCRELS_FTR2 or '') + DOCRELS_TAIL
            p = os.path.join(OUTDIR, f'ef_{case}_{x:04d}.docx')
            with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', ct)
                z.writestr('_rels/.rels', RELS)
                z.writestr('word/_rels/document.xml.rels', dr)
                z.writestr('word/document.xml', doc)
                z.writestr('word/styles.xml', STYLES)
                z.writestr('word/settings.xml', settings(FLAG[case]))
                for k, v in parts.items():
                    z.writestr(k, v)
            n += 1
    print(f'generated {n} -> {os.path.abspath(OUTDIR)}')


def readout(pdf):
    """-> per page dict: filler counts by prefix + which footer ink is drawn."""
    import fitz
    d = fitz.open(pdf)
    out = []
    for pg in d:
        c = {'L': 0, 'A': 0, 'B': 0}
        ink = []
        for blk in pg.get_text('dict')['blocks']:
            if blk.get('type') != 0:
                continue
            for ln in blk['lines']:
                t = ''.join(s['text'] for s in ln['spans'])
                m = re.match(r'^([LAB])\d{3}\b', t)
                if m:
                    c[m.group(1)] += 1
                for tag in ('FTRINK', 'FTRONE', 'FTRTWO'):
                    if tag in t and tag not in ink:
                        ink.append(tag)
        out.append({'n': c['L'] + c['A'] + c['B'], 'L': c['L'], 'A': c['A'],
                    'B': c['B'], 'ink': ink})
    d.close()
    return out


def measure():
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for f in sorted(glob.glob(os.path.join(OUTDIR, 'ef_*.docx'))):
            pdf = f[:-5] + '.pdf'
            if not os.path.exists(pdf):
                d = word.Documents.Open(os.path.abspath(f), ReadOnly=True)
                d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
                d.Close(False)
            res[os.path.basename(f)[:-5]] = readout(pdf)
            print('  measured', os.path.basename(f), flush=True)
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUTDIR, '_measure.json'), 'w'), indent=1)
    report(res)


def report(res=None):
    if res is None:
        res = json.load(open(os.path.join(OUTDIR, '_measure.json')))
    margin_cbot = PAGE_H - BOTTOM_TW / 20.0
    n_margin = int((margin_cbot - 72) // ADV)
    print('\n=== geometry ===')
    print(f'  pageH {PAGE_H}  bottom_margin {BOTTOM_TW/20:.2f}  footer_dist {FOOTER_TW/20:.2f}')
    print(f'  NO-RESERVE model : cbot = {margin_cbot:.2f} -> page-1 capacity n = {n_margin}')
    print('  RESERVE   model : cbot = 841.9 - 64.65 - stack')
    xs = sorted({int(k.rsplit('_', 1)[1]) for k in res})
    print('\n=== predictions ===')
    for x in xs:
        cb = PAGE_H - FOOTER_TW / 20.0 - x / 20.0
        print(f'  stack={x/20:6.2f}pt -> reserve cbot {cb:7.2f} n={int((cb-72)//ADV):3d}'
              f'   | no-reserve cbot {margin_cbot:.2f} n={n_margin}')
    def cell(pg):
        if pg is None:
            return '  ---- '
        tag = {'FTRINK': 'I', 'FTRONE': '1', 'FTRTWO': '2'}
        s = ''.join(tag[t] for t in pg['ink']) or '.'
        return f'{pg["n"]:4d}{s:<3}'

    print('\n=== MEASURED (per page: n_filler + footer ink) ===')
    print('   ink: I=FTRINK  1=FTRONE(sec1 default)  2=FTRTWO(sec2 even)  .=none')
    print('case          ' + ''.join(f'  stack={x/20:6.2f}pt        ' for x in xs))
    for case in CASES:
        row = f'{case:<14}'
        for x in xs:
            pgs = res.get(f'ef_{case}_{x:04d}')
            row += '  ' + ''.join(cell(pgs[i] if pgs and i < len(pgs) else None)
                                  for i in range(3))
        print(row)

    print('\n=== VERDICT 1: does an INERT even ref reserve? (page-1 capacity) ===')
    for x in xs:
        g = lambda c: (res.get(f'ef_{c}_{x:04d}') or [{}])[0].get('n')
        cn, eo, dn = g('ctrl_none'), g('even_only'), g('default_only')
        if None in (cn, eo, dn):
            continue
        v = ('NO RESERVATION (even_only == ctrl_none)' if eo == cn != dn else
             'RESERVES (even_only == default_only)' if eo == dn != cn else
             'INCONCLUSIVE / knob dead' if cn == dn else '???')
        print(f'  stack={x/20:6.2f}pt  ctrl_none={cn}  even_only={eo}  default_only={dn}  -> {v}')

    print('\n=== VERDICT 2: header/footer inheritance is PER TYPE? (inherit_even) ===')
    for x in xs:
        pgs = res.get(f'ef_inherit_even_{x:04d}')
        if not pgs:
            continue
        sec2 = next((p for p in pgs if p['B'] and not p['A']), None)
        if not sec2:
            print(f'  stack={x/20:6.2f}pt  no pure-sec2 page found'); continue
        cb = PAGE_H - FOOTER_TW / 20.0 - x / 20.0
        exp_push = int((cb - 72) // ADV)
        v = ('PER-TYPE INHERIT (sec1 default drawn+reserved on sec2)'
             if 'FTRONE' in sec2['ink'] else
             'even ref honoured (falsifies verdict 1!)' if 'FTRTWO' in sec2['ink'] else
             'NO footer on sec2 (all-or-nothing + even inert)')
        print(f'  stack={x/20:6.2f}pt  sec2 first page: n={sec2["n"]} ink={sec2["ink"] or ["none"]}'
              f'  (push predicts {exp_push}, no-reserve {n_margin})  -> {v}')


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == 'gen':
        gen([int(v) for v in a[1:]] or [400, 1200, 2400])
    elif a and a[0] == 'measure':
        measure()
    elif a and a[0] == 'report':
        report()
    else:
        print(__doc__)
