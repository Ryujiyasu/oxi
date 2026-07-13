"""Justified space-shrink ALLOW boundary — cs-profile sweep.

ukframework 'Board' line render-truth: Word shrinks a NON-LAST justified
line's spaces ~15-18% (proportional across baked cs 19..49tw) to fit the
next word, yet WRAPS 16 other lines at <=16% — no flat cap fits
(OXI_S799_CAP 0.15 inert / 0.16 trades +2 for -16). Hypothesis: the
allowance is a function of the line's BAKED-CS profile (e.g. a fraction
of the cs sum), not of the total space width.

Method: a ~40-word justified paragraph (Calibri 11 — the framework
substitution target), EVERY space its own run carrying w:spacing=CS.
Right margin swept; readout = line 1's LAST word. The k->k-1 flip
happens where needed_shrink == allow, so
    allow_pt = W_natural_fit(k) - W_flip(k)
with W_natural_fit computed from metrics (word widths + (em+cs) spaces).
Variants: cs 0 / 19 / 46 / MIX (19,19,20 then 46.. = the Board profile).

Usage:
  python _pb_cs3_gen.py gen [fine:LO:HI:STEP:CFG]
  python _pb_cs3_gen.py measure [pattern]
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_cs3")

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

R = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/>'
WORDS = [f'word{i:02d}x' for i in range(40)]


def cs_at(cfg, i):
    if cfg == 'c0':
        return 0
    if cfg == 'c19':
        return 19
    if cfg == 'c46':
        return 46
    if cfg == 'cmix':
        return (19, 19, 20)[i] if i < 3 else 46
    raise KeyError(cfg)


def build(right_tw, cfg):
    runs = []
    for i, w in enumerate(WORDS):
        if i:
            v = cs_at(cfg, i - 1)
            sp = f'<w:spacing w:val="{v}"/>' if v else ''
            runs.append(f'<w:r><w:rPr>{R}{sp}</w:rPr><w:t xml:space="preserve"> </w:t></w:r>')
        runs.append(f'<w:r><w:rPr>{R}</w:rPr><w:t>{w}</w:t></w:r>')
    body = (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:jc w:val="both"/><w:rPr>{R}</w:rPr></w:pPr>{"".join(runs)}</w:p>')
    body += (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1440" w:right="{right_tw}" w:bottom="1440" '
             f'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


CFGS = ('c0', 'c19', 'c46', 'cmix')
CASES = [(c, r) for c in CFGS for r in range(600, 1601, 40)]


def gen(cases=None):
    os.makedirs(OUTDIR, exist_ok=True)
    for c, r in (cases or CASES):
        with zipfile.ZipFile(os.path.join(OUTDIR, f'pc3_{c}_{r:04d}.docx'), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/document.xml', build(r, c))
    print('generated', len(cases or CASES), 'docs in', OUTDIR)


def measure(pat='pc3_*'):
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
            l1 = None
            for blk in d[0].get_text('dict')['blocks']:
                if blk.get('type') != 0:
                    continue
                for ln in blk['lines']:
                    t = ''.join(s['text'] for s in ln['spans']).strip()
                    if t.startswith('word00x') and l1 is None:
                        l1 = t
            base = os.path.basename(f)[:-5]
            r = int(base.rsplit('_', 1)[-1])
            last = l1.split()[-1] if l1 else '?'
            print(f'{base}: cright={595.3 - r / 20.0:.1f} l1last={last}')
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        spec = sys.argv[2] if len(sys.argv) > 2 else 'coarse'
        if spec == 'coarse':
            gen()
        else:
            _, lo, hi, step, c = spec.split(':')
            gen([(c, r) for r in range(int(lo), int(hi) + 1, int(step))])
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'pc3_*')
