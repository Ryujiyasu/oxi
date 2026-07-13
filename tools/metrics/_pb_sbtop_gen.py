"""Space-before at page top — natural vs manual break derivation.

uklocalspending p36 render-truth: 'Annex I' (Heading1, inherited
before=240) lands at a NATURAL page break and Word renders it AT the top
margin (ink 72.77 ~ margin 72) — space-before suppressed. Oxi applies
the 12pt (84.2). This sweep pins the rule with controls:

  N filler paragraphs (Arial 11, before=0 after=0, single) then a TARGET
  paragraph with before=240. N swept so the target crosses the p1/p2
  boundary naturally. Variants:
    nat    : natural flow only
    brk    : explicit <w:br w:type="page"/> at the end of the last filler
    pbb    : target carries <w:pageBreakBefore/>
    natsb0 : natural, target before=0 (control — baseline top y)
  Readout: target's page + ink y (PDF). suppressed => y == natsb0's y.

Usage:
  python _pb_sbtop_gen.py gen
  python _pb_sbtop_gen.py measure [pattern]
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_sbtop")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
FILLER = 'Filler paragraph text line for the page fill sweep.'


def build(n_fill, variant):
    paras = []
    for i in range(n_fill):
        brk = ''
        if variant == 'brk' and i == n_fill - 1:
            brk = '<w:r><w:br w:type="page"/></w:r>'
        paras.append(
            f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{R}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R}</w:rPr><w:t>{FILLER} {i:02d}</w:t></w:r>{brk}</w:p>')
    sb = 0 if variant == 'natsb0' else 240
    pbb = '<w:pageBreakBefore/>' if variant == 'pbb' else ''
    paras.append(
        f'<w:p><w:pPr>{pbb}<w:spacing w:before="{sb}" w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:rPr>{R}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{R}</w:rPr><w:t>TARGETLINE space before probe</w:t></w:r></w:p>')
    body = ''.join(paras)
    body += ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
             'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


# A4 portrait content height ~ 697.6pt, Arial 11 line ~12.65 -> ~55 lines/page.
# Sweep n_fill 52..57 for the natural variants; brk/pbb need just one point.
CASES = ([('nat', n) for n in range(52, 58)] +
         [('natsb0', n) for n in range(52, 58)] +
         [('brk', 30), ('pbb', 30)])


def name(v, n):
    return f'pbs_{v}_{n:02d}.docx'


def gen(cases=None):
    os.makedirs(OUTDIR, exist_ok=True)
    for v, n in (cases or CASES):
        with zipfile.ZipFile(os.path.join(OUTDIR, name(v, n)), 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/document.xml', build(n, v))
    print('generated', len(cases or CASES), 'docs in', OUTDIR)


def measure(pat='pbs_*'):
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
                        t = ''.join(s['text'] for s in ln['spans'])
                        if 'TARGETLINE' in t:
                            loc = (pi + 1, round(ln['bbox'][1], 2))
                if loc:
                    break
            base = os.path.basename(f)[:-5]
            print(f'{base}: target page={loc[0] if loc else "?"} y={loc[1] if loc else "?"}')
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        gen()
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'pbs_*')
