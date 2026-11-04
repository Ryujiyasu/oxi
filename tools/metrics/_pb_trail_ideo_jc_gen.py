# -*- coding: utf-8 -*-
"""Right / centre alignment: does a line's TRAILING «　» (U+3000) count toward
its width, as a trailing ASCII space does not?

blind-J ja/reports__7a8daf36 paragraph 13 (jc right): [inline chart]«　　» « »
-- Word sets the chart flush with the right margin (x 235.9..560.0), Oxi 28pt
short (the two U+3000 and the space counted).

Faithful slice of that document (styles, theme, settings, sectPr): one
paragraph per arm, ＭＳ 明朝 12, jc right / center, text «あいう» + a trailing
tail. Readout: x of «あ» and the right edge of «う» against the margin.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/reports/7a8daf3659417f2a.docx'
OUT = 'tests/fixtures/trail_ideo_jc'; os.makedirs(OUT, exist_ok=True)
TAILS = {'none': '', 'ideo2': '　　', 'ideo2sp': '　　 ', 'sp2': '  ', 'ideo1': '　'}
F = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="24"/>'


def body():
    out = ''
    for jc in ('right', 'center'):
        for k, tail in TAILS.items():
            out += (f'<w:p><w:pPr><w:jc w:val="{jc}"/><w:rPr>{F}</w:rPr></w:pPr>'
                    f'<w:r><w:rPr>{F}</w:rPr><w:t xml:space="preserve">あいう{tail}</w:t></w:r></w:p>')
    return out


def build():
    zin = zipfile.ZipFile(SRC)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    at = os.path.abspath(os.path.join(OUT, 'arms.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, data in items:
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                b0 = x.index('<w:body>') + len('<w:body>')
                sect = re.search(r'<w:sectPr[ >].*?</w:sectPr>\s*</w:body>', x, re.S).group(0)
                x = x[:b0] + body() + sect + x[x.index('</w:body>') + len('</w:body>'):]
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
    at = build()
    if '--build' in sys.argv:
        sys.exit()
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        d = app.Documents.Open(at, ReadOnly=True)
        try:
            d.ExportAsFixedFormat(at[:-5] + '.pdf', 17)
        finally:
            d.Close(False)
    finally:
        app.Quit()
    P = pymupdf.open(at[:-5] + '.pdf')[0]
    rows = []
    for b in P.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            for s in l['spans']:
                if 'あいう' in s['text']:
                    rows.append((round(s['origin'][1], 1), round(s['bbox'][0], 2), round(s['bbox'][2], 2), repr(s['text'])))
    names = [f'{jc}_{k}' for jc in ('right', 'center') for k in TAILS]
    for n, r in zip(names, sorted(rows)):
        print(f'{n:16} x0 {r[1]:7.2f}  x1 {r[2]:7.2f}  {r[3]}')
