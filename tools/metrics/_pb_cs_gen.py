"""Baked space-cs (w:spacing on space runs) vs Word's justified line break.

ukframework bullets carry per-space w:spacing (a PDF-converter's baked
justification, cs 19-49tw). Render-truth is contradictory at first sight:
one line fits content whose cs-inclusive width exceeds the column (Word
ignores/shrinks the cs), yet dropping ALL space-cs (S812 v1) over-fits 16
other lines. This sweep pins the rule:

  'Alpha bravo ... oscar mandated' with EVERY space as its own run
  carrying w:spacing val=V; V in {0, 20, 40} x jc {both, left};
  right margin swept. The 1->2 line flip margin vs V gives the fit rule:
    flip independent of V  -> cs fully ignored in the break
    flip shifts by n*V     -> cs fully counted
    partial                -> shrink cap model

Usage:
  python _pb_cs_gen.py gen [fine:LO:HI:STEP:CFG]
  python _pb_cs_gen.py measure [pattern]
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_cs")

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
RH = '<w:rFonts w:ascii="Humnst777 BT" w:hAnsi="Humnst777 BT"/><w:sz w:val="22"/>'
WORDS = ('Alpha bravo charlie delta echo foxtrot golf hotel india juliet '
         'kilo lima mike november oscar mandated').split()


def build(right_tw, cs, jc, rpr=None):
    r = rpr or R
    runs = []
    for i, w in enumerate(WORDS):
        if i:
            sp = f'<w:spacing w:val="{cs}"/>' if cs else ''
            runs.append(f'<w:r><w:rPr>{r}{sp}</w:rPr><w:t xml:space="preserve"> </w:t></w:r>')
        runs.append(f'<w:r><w:rPr>{r}</w:rPr><w:t>{w}</w:t></w:r>')
    body = (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:jc w:val="{jc}"/><w:rPr>{R}</w:rPr></w:pPr>{"".join(runs)}</w:p>')
    body += (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1440" w:right="{right_tw}" w:bottom="1440" '
             f'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


# cfg tag -> (cs, jc)
CFGS = {
    'v0both': (0, 'both'),
    'v20both': (20, 'both'),
    'v40both': (40, 'both'),
    'v40left': (40, 'left'),
    # substituted-font variants (Humnst777 BT is not installed; Word substitutes)
    'vh0both': (0, 'both', 'H'),
    'vh40both': (40, 'both', 'H'),
}
CASES = [(c, r) for c in CFGS for r in range(600, 1801, 20)]


def name(c, r):
    return f'pbc_{c}_{r}.docx'


def gen(cases=None):
    os.makedirs(OUTDIR, exist_ok=True)
    for c, r in (cases or CASES):
        cfg = CFGS[c]
        cs, jc = cfg[0], cfg[1]
        rpr = RH if len(cfg) > 2 else None
        with zipfile.ZipFile(os.path.join(OUTDIR, name(c, r)), 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/document.xml', build(r, cs, jc, rpr))
    print('generated', len(cases or CASES), 'docs in', OUTDIR)


def measure(pat='pbc_*'):
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
            n = 0
            for blk in d[0].get_text('dict')['blocks']:
                if blk.get('type') != 0:
                    continue
                for ln in blk['lines']:
                    if ''.join(s['text'] for s in ln['spans']).strip():
                        n += 1
            base = os.path.basename(f)[:-5]
            r = int(base.rsplit('_', 1)[-1])
            print(f'{base}: lines={n} content_right={595.3 - r / 20.0:.1f}')
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        if len(sys.argv) > 2:
            _, lo, hi, step, c = sys.argv[2].split(':')
            gen([(c, r) for r in range(int(lo), int(hi) + 1, int(step))])
        else:
            gen()
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'pbc_*')
