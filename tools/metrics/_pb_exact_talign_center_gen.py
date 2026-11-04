# -*- coding: utf-8 -*-
"""Where does Word put the baseline of an EXACT line whose paragraph says
textAlignment="center"?

blind-J ja/administrative__208244a3: Normal = line 482 exact (24.1pt) +
textAlignment center, ＭＳ 明朝 10.5 (run face «Mincho»), linesAndChars grid.
Word's first baseline is 15.89 below the line top; Oxi uses the exact-line
law (0.8 x L = 19.28) and every line of page 1 sits 3.3pt low, which pushes
the page's last line onto page 2.

Faithful slice of that document (styles, settings, sectPr): pairs of
paragraphs «あいう» under line L exact (18 / 24.1 / 30) x fs (10.5 / 14) x
textAlignment (center / auto / baseline, set in pPr). Readout: baseline of the
first line below the top margin.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/administrative/208244a32d2bd18e.docx'
OUT = 'tests/fixtures/exact_talign_center'; os.makedirs(OUT, exist_ok=True)
F = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
ARMS = [(L, fs, ta) for L in (360, 482, 600) for fs in (21, 28) for ta in ('center', 'auto', 'baseline')]


def body(L, fs, ta):
    p = (f'<w:p><w:pPr><w:spacing w:line="{L}" w:lineRule="exact"/><w:textAlignment w:val="{ta}"/>'
         f'<w:rPr>{F}<w:sz w:val="{fs}"/></w:rPr></w:pPr>'
         f'<w:r><w:rPr>{F}<w:sz w:val="{fs}"/></w:rPr><w:t>あいう</w:t></w:r></w:p>')
    return p


def build(name, L, fs, ta):
    zin = zipfile.ZipFile(SRC)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, data in items:
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                b0 = x.index('<w:body>') + len('<w:body>')
                sect = re.search(r'<w:sectPr[ >].*?</w:sectPr>\s*</w:body>', x, re.S).group(0)
                x = x[:b0] + body(L, fs, ta) + sect + x[x.index('</w:body>') + len('</w:body>'):]
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
    paths = {f'L{L}_fs{fs}_{ta}': build(f'L{L}_fs{fs}_{ta}', L, fs, ta) for L, fs, ta in ARMS}
    if '--build' in sys.argv:
        sys.exit()
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for n, at in paths.items():
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(at[:-5] + '.pdf', 17)
            finally:
                d.Close(False)
            P = pymupdf.open(at[:-5] + '.pdf')[0]
            y = next(s['origin'][1] for b in P.get_text('dict')['blocks'] for l in b.get('lines', [])
                     for s in l['spans'] if 'あいう' in s['text'])
            print(f'{n:22} baseline - top {y - 1247 / 20:.2f}', flush=True)
    finally:
        app.Quit()
