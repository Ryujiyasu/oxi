# -*- coding: utf-8 -*-
"""Footer TEXT-line height derivation (the S886 open piece).

legal__0001482d wants its Arial-10/8 footer text lines at hhea (11.5/9.2:
stack window (31.7, 34.0] from body ink 630.5 + the wp37 keepNext flip);
usnyserda's S828-S835 zero-drift derivation pins its TNR-10 footer text
line at the 10.5 estimate. Same fonts metrically (hhea sum 2355/2048) —
the two real docs CONTRADICT, so this controlled probe measures Word
directly: ONE footer paragraph per config, geometry always STACK-BOUND
(footer_dist == bottom_margin == 1440tw -> reserved = 72 + stack).

  keep TARGET on p1 iff 72 + K*12.6489 + X/20 + 12.6489 <= 841.9 - 72 - stack
  => stack = 697.9 - (K+1)*12.6489 - X_flip/20   (0.2pt steps; 697.9 = 841.9-144)

Configs (fillers Arial 11 direct spacing 0; footer paras direct spacing 0,
no styles involved):
  fT10 : footer = one TEXT para Arial 10 (sz20)       K=52
  fT8  : footer = one TEXT para Arial 8  (sz16)       K=52
  fE10 : footer = one EMPTY para (mark rPr sz20)      K=52  (harness check:
         both engines already agree on hhea 11.499 here)
  fLGL : legal replica = empty(sz20, pBdr bottom sz4 space1)
         + text(sz20) + text(sz16)                    K=50

Usage:
  python _pb_ftext_gen.py gen                  (coarse 20tw steps)
  python _pb_ftext_gen.py gen fine:LO:HI:STEP:CFG
  python _pb_ftext_gen.py measure [pattern]
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_ftext")

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
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0"/></w:pPr>'
          '</w:style></w:styles>')

R11 = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
R10 = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/>'
R8 = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/>'
SP0 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
ADV = 12.6489

FOOTERS = {
    'fT10': (f'<w:p><w:pPr>{SP0}<w:rPr>{R10}</w:rPr></w:pPr>'
             f'<w:r><w:rPr>{R10}</w:rPr><w:t>Footer text line</w:t></w:r></w:p>'),
    'fT8': (f'<w:p><w:pPr>{SP0}<w:rPr>{R8}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R8}</w:rPr><w:t>Footer text line</w:t></w:r></w:p>'),
    'fE10': f'<w:p><w:pPr>{SP0}<w:rPr>{R10}</w:rPr></w:pPr></w:p>',
    'fLGL': (f'<w:p><w:pPr>{SP0}'
             '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>'
             f'<w:rPr>{R10}</w:rPr></w:pPr></w:p>'
             f'<w:p><w:pPr>{SP0}<w:rPr>{R10}</w:rPr></w:pPr>'
             f'<w:r><w:rPr>{R10}</w:rPr><w:t>As at 31 Mar 2025</w:t></w:r></w:p>'
             f'<w:p><w:pPr>{SP0}<w:rPr>{R8}</w:rPr></w:pPr>'
             f'<w:r><w:rPr>{R8}</w:rPr><w:t>Published on www.example.gov.au</w:t></w:r></w:p>'),
    # nyserda decomposition: TEXT + TRAILING EMPTY (same para props),
    # with/without a pBdr TOP on both (style-level border = merged group).
    'fTE10': (f'<w:p><w:pPr>{SP0}<w:rPr>{R10}</w:rPr></w:pPr>'
              f'<w:r><w:rPr>{R10}</w:rPr><w:t>34</w:t></w:r></w:p>'
              f'<w:p><w:pPr>{SP0}<w:rPr>{R10}</w:rPr></w:pPr></w:p>'),
    'fNY': ('<w:p><w:pPr>'
            '<w:pBdr><w:top w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr>'
            f'{SP0}<w:rPr>{R10}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R10}</w:rPr><w:t>34</w:t></w:r></w:p>'
            '<w:p><w:pPr>'
            '<w:pBdr><w:top w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr>'
            f'{SP0}<w:rPr>{R10}</w:rPr></w:pPr></w:p>'),
    # fNY + style-ish before=240 on both paras (the nyserda spacing) —
    # discriminates border-space-in-spacing (stack 48.5) vs additive (49.5).
    'fNYS': ('<w:p><w:pPr>'
             '<w:pBdr><w:top w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr>'
             '<w:spacing w:before="240" w:after="0" w:line="240" w:lineRule="auto"/>'
             f'<w:rPr>{R10}</w:rPr></w:pPr>'
             f'<w:r><w:rPr>{R10}</w:rPr><w:t>34</w:t></w:r></w:p>'
             '<w:p><w:pPr>'
             '<w:pBdr><w:top w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr>'
             '<w:spacing w:before="240" w:after="0" w:line="240" w:lineRule="auto"/>'
             f'<w:rPr>{R10}</w:rPr></w:pPr></w:p>'),
}
CFG_K = {'fT10': 52, 'fT8': 52, 'fE10': 52, 'fLGL': 50, 'fTE10': 51, 'fNY': 51,
         'fNYS': 49}


def body(k, spacer_tw):
    paras = []
    for i in range(k):
        paras.append(
            f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R11}</w:rPr><w:t>Item {i:02d} alpha beta gamma delta.</w:t></w:r></w:p>')
    paras.append(
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{spacer_tw}" w:lineRule="exact"/>'
        f'<w:rPr>{R11}</w:rPr></w:pPr></w:p>')
    paras.append(
        f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{R11}</w:rPr><w:t>TARGETLINE omega.</w:t></w:r></w:p>')
    return ''.join(paras)


def build(cfg, spacer_tw):
    b = body(CFG_K[cfg], spacer_tw)
    b += ('<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
          '<w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
          'w:left="1440" w:header="709" w:footer="1440" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{b}</w:body></w:document>')


def gen(cases):
    os.makedirs(OUTDIR, exist_ok=True)
    for cfg, x in cases:
        p = os.path.join(OUTDIR, f'pft_{cfg}_{x:04d}.docx')
        with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', build(cfg, x))
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/footer1.xml',
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       f'<w:ftr {W_NS}>{FOOTERS[cfg]}</w:ftr>')
    print('generated', len(cases), 'docs in', os.path.abspath(OUTDIR))


def measure(pat='pft_*'):
    import glob
    import win32com.client, fitz
    word = win32com.client.Dispatch('Word.Application')
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
                            loc = (pi + 1, round(ln['bbox'][1], 2))
                if loc:
                    break
            d.close()
            base = os.path.basename(f)[:-5]
            cfg, xs = base.split('_')[1], base.split('_')[2]
            x = int(xs)
            k = CFG_K[cfg]
            # stack implied when THIS X is the last keep:
            stack = 697.9 - (k + 1) * ADV - x / 20.0  # 841.9 - 2*72 (S889 bug: was 697.25)
            print(f'{base}: page={loc[0] if loc else "?"} '
                  f'stack_if_lastkeep={stack:.2f}')
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        spec = sys.argv[2] if len(sys.argv) > 2 else 'coarse'
        cases = []
        if spec == 'coarse':
            for cfg in ('fT10', 'fT8', 'fE10'):
                for x in range(200, 420, 20):
                    cases.append((cfg, x))
            for x in range(300, 520, 20):
                cases.append(('fLGL', x))
        else:
            _, lo, hi, step, cfg = spec.split(':')
            for x in range(int(lo), int(hi) + 1, int(step)):
                cases.append((cfg, x))
        gen(cases)
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'pft_*')
