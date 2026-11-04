# -*- coding: utf-8 -*-
"""Typed docGrid: does a larger WHITESPACE run set where Word centres the line's
box in its grid cell(s)?

ja/administrative__05e3143c: the body paragraph after the boxed heading opens
with a 16pt «　» (HG丸ｺﾞｼｯｸM-PRO) before 12pt text; Word's baseline sits
13.32 below the line top (the 12pt box centred in the 18pt cell), Oxi's 14.75
(the 16pt box centred).

Faithful slice of that document (styles, theme, settings, sectPr; docGrid lines
360): «あいう» (style 2 at 12pt) then X, X = 12pt text after a 16pt «　» / a
16pt ASCII space / a 16pt «あ» (control) / with a trailing 16pt «　» / alone.
Readout: the baseline gap «あいう» -> X's text (18 = same centring as a
12pt line).
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/administrative/05e3143c61176854.docx'
OUT = 'tests/fixtures/grid_space_center'; os.makedirs(OUT, exist_ok=True)
PPR = '<w:pPr><w:pStyle w:val="2"/><w:rPr><w:b w:val="0"/><w:bCs w:val="0"/><w:w w:val="100"/><w:sz w:val="24"/></w:rPr></w:pPr>'


def r(t, sz):
    return (f'<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:b w:val="0"/><w:bCs w:val="0"/><w:w w:val="100"/>'
            f'<w:sz w:val="{sz}"/></w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>')


ARMS = {
    'alone': r('かきくけこ', 24),
    'ideo16': r('　', 32) + r('かきくけこ', 24),
    'ascii16': r(' ', 32) + r('かきくけこ', 24),
    'glyph16': r('さ', 32) + r('かきくけこ', 24),
    'trail16': r('かきくけこ', 24) + r('　', 32),
    'ideo20': r('　', 40) + r('かきくけこ', 24),
}


def build(name, x_runs):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                b0 = x.index('<w:body>') + len('<w:body>')
                sect = re.search(r'<w:sectPr[ >].*?</w:sectPr>\s*</w:body>', x, re.S).group(0)
                body = f'<w:p>{PPR}{r("あいう", 24)}</w:p><w:p>{PPR}{x_runs}</w:p>'
                x = x[:b0] + body + sect + x[x.index('</w:body>') + len('</w:body>'):]
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
    paths = {n: build(n, m) for n, m in ARMS.items()}
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
            ys = {}
            for b in P.get_text('dict')['blocks']:
                for l in b.get('lines', []):
                    for s in l['spans']:
                        t = s['text'].strip()
                        if t.startswith('あいう'):
                            ys['a'] = s['origin'][1]
                        elif 'かきくけこ' in t:
                            ys['x'] = s['origin'][1]
            print(f'{n:8} gap {ys["x"] - ys["a"]:.2f}', flush=True)
    finally:
        app.Quit()
