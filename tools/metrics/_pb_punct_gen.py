"""Latin trailing-punctuation hang (overflowPunct) derivation.

uklocalspending p36 render-truth: Word fits '...are not mandated.' with
'mandated' ending AT the content-right (769.97 ~ 769.9) and the trailing
'.' (3.05pt) HANGING past the margin — the Latin analog of the CJK
burasage. This sweep pins the rule with a control:

  One paragraph 'Alpha bravo ... zulu mandated<P>' (P = '.', ',' or 'x'),
  Arial 11, A4 portrait, RIGHT margin swept in 2tw steps. The flip margin
  where the para goes 1 line -> 2 lines tells whether P's width counts:
    control 'x': flips when 'mandatedx' no longer fits (full width)
    '.' / ',': if Word hangs, flips ~P-width LATER (only 'mandated' must fit)

Usage:
  python _pb_punct_gen.py gen
  python _pb_punct_gen.py measure
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_punct")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '</Types>')

SETTINGS15 = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    '<w:compat><w:compatSetting w:name="compatibilityMode" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')

DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
           '</Relationships>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
TEXT = 'Alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november oscar mandated'


def build(right_tw, punct):
    txt = TEXT + punct
    body = (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{R}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R}</w:rPr><w:t>{txt}</w:t></w:r></w:p>')
    body += (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1440" w:right="{right_tw}" w:bottom="1440" '
             f'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


PUNCTS = {'dot': '.', 'comma': ',', 'ctrl': 'x',
          'dot15': '.', 'ctrl15': 'x'}
# text 'Alpha ... mandated' at Arial 11 is ~ 520pt; A4 width 595.3, left 72 ->
# content must shrink to ~520 to straddle: right margin ~ 595.3-72-520 = 3.3pt?
# too small. Use a longer text? Instead sweep right from 60 to 130pt in 2tw
# steps to cross the boundary wherever it is.
CASES = [(p, r) for p in PUNCTS for r in range(600, 1801, 20)]


def name(p, r):
    return f'pbp_{p}_{r}.docx'


def gen(cases=None):
    os.makedirs(OUTDIR, exist_ok=True)
    cs = cases or CASES
    for p, r in cs:
        with zipfile.ZipFile(os.path.join(OUTDIR, name(p, r)), 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/document.xml', build(r, PUNCTS[p]))
            if p.endswith('15'):
                z.writestr('word/_rels/document.xml.rels', DOCRELS)
                z.writestr('word/settings.xml', SETTINGS15)
    print('generated', len(cs), 'docs in', OUTDIR)


def measure(pat='pbp_*'):
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
            last_end = 0
            for blk in d[0].get_text('dict')['blocks']:
                if blk.get('type') != 0:
                    continue
                for ln in blk['lines']:
                    t = ''.join(s['text'] for s in ln['spans'])
                    if t.strip():
                        n += 1
                        last_end = round(ln['bbox'][2], 2)
            base = os.path.basename(f)[:-5]
            r = int(base.rsplit('_', 1)[-1])
            cr = 595.3 - r / 20.0
            print(f'{base}: lines={n} content_right={cr:.1f}')
    finally:
        word.Quit()


if __name__ == '__main__':
    if sys.argv[1] == 'gen':
        if len(sys.argv) > 2:
            _, lo, hi, step, p = sys.argv[2].split(':')
            gen([(p, r) for r in range(int(lo), int(hi) + 1, int(step))])
        else:
            gen()
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'pbp_*')
