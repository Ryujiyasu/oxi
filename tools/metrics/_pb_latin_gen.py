"""LM0 (no-docGrid) LATIN page-bottom rule derivation — controlled bottom-margin sweep.

The nyserda finding (2026-07-10): Word pushes a 12pt TNR continuation line whose
baseline+13.8 exceeds content_bottom (base 707.26+13.8=721.06 > 720) while Oxi
fits line_top+13.8 <= 720 -> a one-line page-phase shift that tanks per-page SSIM
(0.54 vs Libra 0.95). This sweep pins Word's exact threshold.

Each doc: [Part A: 60 single-line paragraphs "Item i alpha..."] [page break]
[Part B: one long paragraph wrapping ~90 lines]. Letter page, top=1440,
bottom SWEPT. Measurement = Word COM ExportAsFixedFormat -> fitz baselines:
  A: last baseline on page 1 + count           (paragraph-line rule)
  B: last baseline on B's first page + count   (continuation-line rule)

Usage:
  python _pb_latin_gen.py gen [coarse|fine:LO:HI:STEP]
  python _pb_latin_gen.py measure
"""
import os, sys, json, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_latin")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CONTENT_TYPES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>')

def rpr(font, sz):
    return f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz}"/>'

def para_single(i, font, sz):
    r = rpr(font, sz)
    return (f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t>Item {i} alpha beta gamma delta.</w:t></w:r></w:p>')

def para_long(font, sz):
    r = rpr(font, sz)
    sent = ("The contractor shall provide all services in a manner consistent "
            "with the terms of this agreement and applicable law. ") * 60
    return (f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">BSTART {sent}</w:t></w:r></w:p>')

def para_pagebreak():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

# variant: (tag, font, sz_halfpt, grid_xml)
VARIANTS = {
    'tnr12': ('Times New Roman', 24, ''),
    'ari11': ('Arial', 22, ''),
    'cal11': ('Calibri', 22, ''),
    'tnr10': ('Times New Roman', 20, ''),
    'a11g360': ('Arial', 22, '<w:docGrid w:linePitch="360"/>'),
    't12g360': ('Times New Roman', 24, '<w:docGrid w:linePitch="360"/>'),
}

def build(bottom_tw, variant='tnr12'):
    font, sz, grid = VARIANTS[variant]
    body = ''.join(para_single(i + 1, font, sz) for i in range(80))
    body += para_pagebreak()
    body += para_long(font, sz)
    body += (f'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
             f'<w:pgMar w:top="1440" w:right="1800" w:bottom="{bottom_tw}" '
             f'w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>{grid}</w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')

def gen(cases, variant='tnr12'):
    os.makedirs(OUTDIR, exist_ok=True)
    for b in cases:
        p = os.path.join(OUTDIR, f'pbl_{variant}_{b}.docx')
        with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CONTENT_TYPES)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/document.xml', build(b, variant))
    print('generated', len(cases), variant, 'docs in', OUTDIR)

def measure():
    import glob
    import win32com.client
    import fitz
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    out = {}
    pat = sys.argv[2] if len(sys.argv) > 2 else 'pbl_*'
    for p in sorted(glob.glob(os.path.join(OUTDIR, pat + '.docx'))):
        b = os.path.basename(p)[:-5]
        pdf = p[:-5] + '.pdf'
        try:
            doc = word.Documents.Open(os.path.abspath(p), ReadOnly=True)
            doc.ExportAsFixedFormat(os.path.abspath(pdf), 17)
            doc.Close(False)
        except Exception as e:
            print(b, 'COM ERR', str(e)[:60]); continue
        d = fitz.open(pdf)
        # Part A = page 1; Part B starts on the page containing BSTART
        def page_bases(pg):
            bases = []
            for blk in pg.get_text('dict')['blocks']:
                for l in blk.get('lines', []):
                    t = ''.join(s['text'] for s in l['spans'])
                    if t.strip():
                        bases.append((round(l['spans'][0]['origin'][1], 2), t[:12]))
            bases.sort()
            return bases
        a = page_bases(d[0])
        bpage = None
        for pi in range(1, len(d)):
            txt = d[pi].get_text()
            if 'BSTART' in txt:
                bpage = pi; break
        bres = page_bases(d[bpage]) if bpage is not None else []
        out[b] = {
            'a_count': len(a), 'a_last_base': a[-1][0] if a else None,
            'b_count': len(bres), 'b_last_base': bres[-1][0] if bres else None,
            'bottom_pt': round(15840 / 20 / 20, 2),
        }
        btw = int(str(b).rsplit('_', 1)[-1])
        content_bottom = 792.0 - btw / 20.0
        print(f'bottom={b}tw ({btw/20:.1f}pt, cbot={content_bottom:.1f}): '
              f'A n={len(a)} last_base={a[-1][0] if a else None} gap={content_bottom - a[-1][0]:.2f} | '
              f'B n={len(bres)} last_base={bres[-1][0] if bres else None} gap={content_bottom - bres[-1][0]:.2f}')
    word.Quit()
    json.dump(out, open(os.path.join(OUTDIR, '_results.json'), 'w'))

if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        spec = sys.argv[2] if len(sys.argv) > 2 else 'coarse'
        variant = sys.argv[3] if len(sys.argv) > 3 else 'tnr12'
        if spec == 'coarse':
            cases = list(range(1300, 1601, 20))
        else:
            _, lo, hi, step = spec.split(':')
            cases = list(range(int(lo), int(hi) + 1, int(step)))
        gen(cases, variant)
    else:
        measure()
