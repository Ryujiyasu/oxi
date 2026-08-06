"""Footer reservation for a 3-PARAGRAPH footer whose FIRST para is empty+pBdr and
whose LAST para is a SMALLER font (legal__0014c86f geometry).

legal__0014c86f (blindB50) loses 3 paragraphs because Oxi's body bottom is ~8pt
higher than Word's.  Measured on its p14 (Word PDF, box tops derived from the
span baselines with Arial hhea):

    last body line  box 625.39..639.19            (TNR 12, 13.799)
    footer p0       box 639.44..650.94  (empty, Arial 10, hhea 11.499)
    p0 bottom pBdr  651.94..652.42      (= p0 box bottom + space 1pt, sz4)
    footer p1       box 652.40..663.90  (Arial 10 text)
    footer p2       box 663.88..673.08  (Arial  8 text, hhea 9.199)
    => footer top 639.44, stack 33.64, footer BOTTOM 673.08

The doc has pgMar bottom == footer == 3542tw (177.1pt), so
`pageH - footer_dist` = 664.9 — yet the footer BOTTOM sits 8.18pt BELOW that and
its TOP is 25.46 above it.  Oxi's S806 model (reserved = max(bottom_margin,
footer_dist + stack) = 210.8 -> cbot 631.2) is therefore ~8pt too generous here,
even though its STACK (33.7) matches Word's 33.64 exactly.

So the open question is not the stack but the ANCHOR: how much of a multi-line
footer sits below `pageH - footer_dist`?  Candidates:
  H1  reserved = footer_dist + stack                        (S806, Oxi today)
  H2  reserved = footer_dist + stack - last_para_height     (last para hangs)
  H3  reserved = footer_dist + stack - (stack - first_two)  (i.e. only the paras
      above the last one count)
  H4  the pBdr space/width is not reserved

Design = the _pb_fstack_gen exact-spacer ladder: K filler lines (TNR 12, spacing
0, single = 13.799) + ONE empty spacer with line=X lineRule=exact + a TARGET
line.  X swept finely; the p1->p2 flip of TARGET pins cbot to STEP/20 pt:

    keep iff  72 + K*ADV + X/20 + ADV <= cbot
    cbot in ( 72 + K*ADV + X_firstpush/20 + ADV - STEP/20 ,
              72 + K*ADV + X_lastkeep /20 + ADV ]

Configs (all A4, top/left/right 1440, header 706, footer == bottom == 3542):
    f0  no footer                       -> control, cbot = 841.9 - 177.1 = 664.8
    f1  1 para  (Arial 10 text)
    f2  2 paras (Arial 10, Arial 8)
    f3  3 paras (empty Arial 10, Arial 10, Arial 8)      = target minus border
    f4  = f3 with the p0 bottom pBdr                     = the target exactly
Increments isolate: f1 = one 10pt line, f2-f1 = the 8pt line, f3-f2 = the empty
para, f4-f3 = the border.  Comparing each against H1..H4 answers the anchor.

Styles are the target's verbatim: docDefaults TNR (no pPrDefault), Normal sz=24,
Footer basedOn Normal + rFonts Arial; the footer paragraphs carry sz in rPr.

Usage:
  python _pb_ftr3_gen.py gen [coarse | fine:LO:HI:STEP:CFG]
  python _pb_ftr3_gen.py measure [pattern]
"""
import os
import sys
import zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_ftr3")

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

CT_FTR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
          '</Types>')

CT_NOFTR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
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

DOCRELS_FTR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
               '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
               '</Relationships>')

DOCRELS_NOFTR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                 '</Relationships>')

# legal__0014c86f styles, verbatim shape
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" '
          'w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:qFormat/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/>'
          '<w:basedOn w:val="Normal"/>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:style>'
          '</w:styles>')

PBDR = ('<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>')


def fpara(sz, text=None, bdr=False):
    ppr = ('<w:pPr><w:pStyle w:val="Footer"/>'
           + (PBDR if bdr else '')
           + f'<w:rPr><w:sz w:val="{sz}"/></w:rPr></w:pPr>')
    run = (f'<w:r><w:rPr><w:sz w:val="{sz}"/></w:rPr><w:t>{text}</w:t></w:r>'
           if text else '')
    return f'<w:p>{ppr}{run}</w:p>'


FOOTERS = {
    'f1': fpara(20, 'Version 00-k0-09 Ceased on 19 Nov 1999'),
    'f2': (fpara(20, 'Version 00-k0-09 Ceased on 19 Nov 1999')
           + fpara(16, 'Extract from www.slp.wa.gov.au, see that website.')),
    'f3': (fpara(20)
           + fpara(20, 'Version 00-k0-09 Ceased on 19 Nov 1999')
           + fpara(16, 'Extract from www.slp.wa.gov.au, see that website.')),
    'f4': (fpara(20, bdr=True)
           + fpara(20, 'Version 00-k0-09 Ceased on 19 Nov 1999')
           + fpara(16, 'Extract from www.slp.wa.gov.au, see that website.')),
}

ADV = 13.7988          # TNR 12 hhea (S805)
BOTTOM_TW = 3542       # == the target's pgMar bottom AND w:footer
K = int(os.environ.get("PB_FTR3_K", "40"))   # filler lines; PB_FTR3_K lowers the band


def body(spacer_tw):
    paras = []
    for i in range(K):
        paras.append(
            '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" '
            'w:lineRule="auto"/></w:pPr>'
            f'<w:r><w:t>Item {i:02d} alpha beta gamma delta epsilon.</w:t></w:r></w:p>')
    paras.append(
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{spacer_tw}" '
        'w:lineRule="exact"/></w:pPr></w:p>')
    paras.append(
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" '
        'w:lineRule="auto"/></w:pPr>'
        '<w:r><w:t>TARGETLINE omega.</w:t></w:r></w:p>')
    return ''.join(paras)


def build(cfg, spacer_tw):
    has_ftr = cfg != 'f0'
    b = body(spacer_tw)
    ref = '<w:footerReference w:type="default" r:id="rId2"/>' if has_ftr else ''
    b += (f'<w:sectPr>{ref}<w:pgSz w:w="11907" w:h="16840" w:code="9"/>'
          f'<w:pgMar w:top="1440" w:right="1440" w:bottom="{BOTTOM_TW}" '
          f'w:left="1440" w:header="706" w:footer="{BOTTOM_TW}" w:gutter="0"/>'
          '</w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{b}</w:body></w:document>')
    return doc, has_ftr


def gen(cases):
    os.makedirs(OUTDIR, exist_ok=True)
    for cfg, x in cases:
        doc, has_ftr = build(cfg, x)
        p = os.path.join(OUTDIR, f'ft3_{cfg}_k{K}_{x:04d}.docx')
        with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT_FTR if has_ftr else CT_NOFTR)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels',
                       DOCRELS_FTR if has_ftr else DOCRELS_NOFTR)
            z.writestr('word/document.xml', doc)
            z.writestr('word/styles.xml', STYLES)
            if has_ftr:
                z.writestr('word/footer1.xml',
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                           f'<w:ftr {W_NS}>{FOOTERS[cfg]}</w:ftr>')
    print('generated', len(cases), 'docs in', OUTDIR)


def measure(pat='ft3_*'):
    import glob
    import fitz
    import win32com.client
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    try:
        for f in sorted(glob.glob(os.path.join(OUTDIR, pat + '.docx'))):
            pdf = f[:-5] + '.pdf'
            if not os.path.exists(pdf):
                doc = word.Documents.Open(os.path.abspath(f), ReadOnly=True)
                doc.ExportAsFixedFormat(os.path.abspath(pdf), 17)
                doc.Close(False)
            d = fitz.open(pdf)
            loc = None
            for pi in range(len(d)):
                for blk in d[pi].get_text('dict')['blocks']:
                    if blk.get('type') != 0:
                        continue
                    for ln in blk['lines']:
                        if 'TARGETLINE' in ''.join(s['text'] for s in ln['spans']):
                            loc = (pi + 1, round(ln['spans'][0]['origin'][1], 2))
                if loc:
                    break
            d.close()
            base = os.path.basename(f)[:-5]
            _, cfg, _k, xs = base.split('_')
            x = int(xs)
            ttop = 72 + K * ADV + x / 20.0
            print(f'{base}: page={loc[0] if loc else "?"} '
                  f'base={loc[1] if loc else "?"} '
                  f'model_top={ttop:7.2f} model_bot={ttop + ADV:7.2f}', flush=True)
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        spec = sys.argv[2] if len(sys.argv) > 2 else 'coarse'
        if spec == 'coarse':
            cases = [(c, x) for c in ('f0', 'f1', 'f2', 'f3', 'f4')
                     for x in range(0, 901, 50)]
            gen(cases)
        else:
            _, lo, hi, step, cfg = spec.split(':')
            gen([(cfg, x) for x in range(int(lo), int(hi) + 1, int(step))])
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'ft3_*')
