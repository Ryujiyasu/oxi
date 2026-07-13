"""Footer-stack DEVICE-precision derivation — exact-spacer sweep.

_pb_footer_gen pinned the S806 stack model only to ~one line pitch (the
bottom-margin sweep can't move the flip in the footer-bound regime; the
line-count readout brackets cbot by 12.65pt). uklocalspending's bundle
blocker needs cbot to ~0.1pt: Word's rendered footer region top ~472.8
vs Oxi's computed 473.2 (stack 85.95 computed vs ~87.05 rendered).

Method: K filler lines (Arial 11, spacing 0, single) + ONE empty spacer
paragraph with line=X lineRule=exact (spacing 0) + a TARGET line. X swept
in fine steps; the p1->p2 flip of TARGET pins cbot:
  keep iff  72 + K*adv + X/20 + hhea_line <= cbot     (S779)
  cbot = 72 + K*adv + X_lastkeep/20 + 12.6489 (+ step/20 bracket)

Configs (A4 portrait, top=1440 bottom=1440 header/footer=709):
  c0 : no footer            -> cbot = 841.9 - 72 = 769.9 (control)
  cE : footer = 1 empty Normal para (before/after 240, Arial 11)
  cS : footer = single Footer-style PAGE para (uklocal styles:
       Footer basedOn Normal + line=260 exact; Normal before/after 240)
  cU : uklocalspending replica footer (Footer PAGE + Footer empty +
       plain <w:p/>)
Increments decompose the stack: cU-cS = the 2 trailing paras; cE = one
auto-line para stack.

Usage:
  python _pb_fstack_gen.py gen coarse            (25tw steps, prediction bands)
  python _pb_fstack_gen.py gen fine:LO:HI:STEP:CFG
  python _pb_fstack_gen.py measure [pattern]
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_fstack")

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

# uklocalspending styles (== _pb_footer_gen STYLES)
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="240" w:after="240"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs>'
          '<w:spacing w:line="260" w:lineRule="exact"/></w:pPr></w:style>'
          '</w:styles>')

PAGE_FLD = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>1</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>')

FOOTERS = {
    'cE': '<w:p/>',
    'cS': ('<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>'
           f'{PAGE_FLD}</w:p>'),
    'cU': ('<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>'
           f'{PAGE_FLD}</w:p>'
           '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>'
           '<w:p/>'),
}

R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
ADV = 12.6489  # Arial 11 hhea (S805)

# per-config filler count K + expected flip X (tw), from stack predictions:
#   c0: cbot 769.9 ; cE: 841.9-35.45-(12+12.65+12)=757.4?? -> measured
#   cS: stack ~ 12+13+12 = 37 -> cbot ~769.45 capped by margin -> ~769?
#   cU: stack ~87 -> cbot ~719.4
# K chosen so the flip lands at X in [200, 900].
CFG_K = {'c0': 53, 'cE': 52, 'cS': 53, 'cU': 49}


def body(k, spacer_tw):
    paras = []
    for i in range(k):
        paras.append(
            f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{R}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R}</w:rPr><w:t>Item {i:02d} alpha beta gamma delta.</w:t></w:r></w:p>')
    paras.append(
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{spacer_tw}" w:lineRule="exact"/>'
        f'<w:rPr>{R}</w:rPr></w:pPr></w:p>')
    paras.append(
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:rPr>{R}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{R}</w:rPr><w:t>TARGETLINE omega.</w:t></w:r></w:p>')
    return ''.join(paras)


def build(cfg, spacer_tw):
    has_ftr = cfg != 'c0'
    b = body(CFG_K[cfg], spacer_tw)
    ftr_ref = '<w:footerReference w:type="default" r:id="rId2"/>' if has_ftr else ''
    b += (f'<w:sectPr>{ftr_ref}<w:pgSz w:w="11906" w:h="16838"/>'
          f'<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
          f'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{b}</w:body></w:document>')
    return doc, has_ftr


def gen(cases):
    os.makedirs(OUTDIR, exist_ok=True)
    for cfg, x in cases:
        doc, has_ftr = build(cfg, x)
        p = os.path.join(OUTDIR, f'pfs_{cfg}_{x:04d}.docx')
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


def measure(pat='pfs_*'):
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
            base = os.path.basename(f)[:-5]
            cfg, xs = base.split('_')[1], base.split('_')[2]
            x = int(xs)
            k = CFG_K[cfg]
            # model: target_top = 72 + k*ADV + x/20 ; cbot >= target_top + ADV while kept
            ttop = 72 + k * ADV + x / 20.0
            print(f'{base}: page={loc[0] if loc else "?"} y={loc[1] if loc else "?"} '
                  f'model_top={ttop:.2f} model_bot={ttop + ADV:.2f}')
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        spec = sys.argv[2] if len(sys.argv) > 2 else 'coarse'
        if spec == 'coarse':
            cases = []
            for cfg in CFG_K:
                cases += [(cfg, x) for x in range(100, 1101, 50)]
            gen(cases)
        else:
            _, lo, hi, step, cfg = spec.split(':')
            gen([(cfg, x) for x in range(int(lo), int(hi) + 1, int(step))])
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'pfs_*')
