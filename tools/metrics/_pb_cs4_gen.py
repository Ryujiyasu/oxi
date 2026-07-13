"""S825 capacity FONT/SIZE discriminator — TNR 12pt c0 sweep (compat-15).

The S825 model (capacity/space = 0.365*em + 0.24*cs) was derived at
Calibri 11 ONLY, where 0.365*em (0.9077) is DEGENERATE with an absolute
constant (~0.91pt/space) and an fs-scaled one (0.0825*fs). nyserda p24
(TNR 12, 13 spaces, needed 12.9 = 0.99/space) shows Word WRAPPING where
the em-fraction model grants 14.2 — excluding the em-fraction. This
sweep at TNR 12 separates all three:
    absolute 0.908/space -> allow(9 sp) = 8.17
    fs-scaled 0.0825*fs  -> allow      = 8.91
    em-fraction 0.365*em -> allow      = 9.86

Usage:
  python _pb_cs4_gen.py gen [fine:LO:HI:STEP]
  python _pb_cs4_gen.py measure
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_cs4")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
           '</Relationships>')

SETTINGS15 = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              f'<w:settings {W_NS} xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
              '<w:compat><w:compatSetting w:name="compatibilityMode" '
              'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
              '</w:settings>')

R = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>'
WORDS = [f'word{i:02d}x' for i in range(40)]


def build(right_tw):
    runs = []
    for i, w in enumerate(WORDS):
        if i:
            runs.append(f'<w:r><w:rPr>{R}</w:rPr><w:t xml:space="preserve"> </w:t></w:r>')
        runs.append(f'<w:r><w:rPr>{R}</w:rPr><w:t>{w}</w:t></w:r>')
    body = (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:jc w:val="both"/><w:rPr>{R}</w:rPr></w:pPr>{"".join(runs)}</w:p>')
    body += (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1440" w:right="{right_tw}" w:bottom="1440" '
             f'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


CASES = list(range(600, 1601, 40))


def gen(cases=None):
    os.makedirs(OUTDIR, exist_ok=True)
    for r in (cases or CASES):
        with zipfile.ZipFile(os.path.join(OUTDIR, f'pc4_t12_{r:04d}.docx'), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/settings.xml', SETTINGS15)
            z.writestr('word/document.xml', build(r))
    print('generated', len(cases or CASES))


def measure(pat='pc4_*'):
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
            print(f'{base}: cright={595.3 - r / 20.0:.1f} l1last={last}', flush=True)
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        spec = sys.argv[2] if len(sys.argv) > 2 else 'coarse'
        if spec == 'coarse':
            gen()
        else:
            _, lo, hi, step = spec.split(':')
            gen(list(range(int(lo), int(hi) + 1, int(step))))
    else:
        measure()
