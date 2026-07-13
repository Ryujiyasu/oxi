"""Justify-shrink enabling condition — MULTI-LINE variant of _pb_cs_gen.

_pb_cs_gen (single-line para = the para's LAST line) proved Word grants
ZERO space shrink at the last-line boundary (flip = full-cs width exact).
ukframework wp15 render-truth shows a NON-LAST justified line shrinking
~6.1pt (of 14.75 baked cs) to fit 'Board'. Hypothesis: justified
NON-LAST lines may compress spaces (down to natural? bounded?); the LAST
line (rendered left-aligned) may not.

Probe: a ~3-line justified para of numbered words, per-space w:spacing V;
margin swept. Readout = LINE 1's last word index + ink end (PDF). Model
predictions per margin:
  cs-counted, no shrink : line1 packs to sum(words + (2.5+V)*gaps) <= W
  full shrink to natural: line1 packs to sum(words + 2.5*gaps) <= W
  partial               : in between.

Usage:
  python _pb_cs2_gen.py gen [fine:LO:HI:STEP:CFG]
  python _pb_cs2_gen.py measure [pattern]
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_cs2")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      '</Types>')

DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
           '</Relationships>')

# ukframework-shaped bullet level (Symbol sz=20 marker, hanging indent)
NUM = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       f'<w:numbering {W_NS}>'
       '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
       '<w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/>'
       '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
       '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/><w:sz w:val="20"/></w:rPr></w:lvl></w:abstractNum>'
       '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
# 36 distinct words, ~3 lines at ~450pt column
WORDS = [f'word{i:02d}x' for i in range(36)]


def build(right_tw, cs, jc, variant=None):
    runs = []
    for i, w in enumerate(WORDS):
        if i:
            v = cs
            if variant == 'mix':
                v = 19 if i % 2 else 46
            sp = f'<w:spacing w:val="{v}"/>' if v else ''
            runs.append(f'<w:r><w:rPr>{R}{sp}</w:rPr><w:t xml:space="preserve"> </w:t></w:r>')
        runs.append(f'<w:r><w:rPr>{R}</w:rPr><w:t>{w}</w:t></w:r>')
    if variant in ('full', 'mix'):
        ppr = (f'<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
               f'<w:spacing w:before="0" w:after="0" w:line="264" w:lineRule="auto"/>'
               f'<w:ind w:right="116"/><w:jc w:val="{jc}"/><w:rPr>{R}</w:rPr></w:pPr>')
    elif variant == 'ind':
        ppr = (f'<w:pPr><w:spacing w:before="0" w:after="0" w:line="264" w:lineRule="auto"/>'
               f'<w:ind w:right="116"/><w:jc w:val="{jc}"/><w:rPr>{R}</w:rPr></w:pPr>')
    else:
        ppr = (f'<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
               f'<w:jc w:val="{jc}"/><w:rPr>{R}</w:rPr></w:pPr>')
    body = f'<w:p>{ppr}{"".join(runs)}</w:p>'
    body += (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1440" w:right="{right_tw}" w:bottom="1440" '
             f'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


CFGS = {
    'm0both': (0, 'both'),
    'm40both': (40, 'both'),
    'm40left': (40, 'left'),
    # framework-faithful pPr variants (cs=46 = the wp15 line's values)
    'f46full': (46, 'both', 'full'),   # numPr + ind right 116 + line=264
    'f46ind': (46, 'both', 'ind'),     # ind right + line=264 only (no numPr)
    'f46mix': (46, 'both', 'mix'),     # full pPr + MIXED cs (alternate 19/46)
}
CASES = [(c, r) for c in CFGS for r in range(600, 1601, 50)]


def name(c, r):
    return f'pb2_{c}_{r}.docx'


def gen(cases=None):
    os.makedirs(OUTDIR, exist_ok=True)
    for c, r in (cases or CASES):
        cfg = CFGS[c]
        cs, jc = cfg[0], cfg[1]
        variant = cfg[2] if len(cfg) > 2 else None
        with zipfile.ZipFile(os.path.join(OUTDIR, name(c, r)), 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/numbering.xml', NUM)
            z.writestr('word/document.xml', build(r, cs, jc, variant))
    print('generated', len(cases or CASES), 'docs in', OUTDIR)


def measure(pat='pb2_*'):
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
            lines = []
            for blk in d[0].get_text('dict')['blocks']:
                if blk.get('type') != 0:
                    continue
                for ln in blk['lines']:
                    t = ''.join(s['text'] for s in ln['spans'])
                    if t.strip():
                        lines.append((round(ln['bbox'][1], 1), t.strip()))
            lines.sort()
            base = os.path.basename(f)[:-5]
            r = int(base.rsplit('_', 1)[-1])
            l1 = lines[0][1] if lines else ''
            last_word = l1.split()[-1] if l1 else ''
            print(f'{base}: cright={595.3 - r / 20.0:.1f} nlines={len(lines)} l1last={last_word}')
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
        measure(sys.argv[2] if len(sys.argv) > 2 else 'pb2_*')
