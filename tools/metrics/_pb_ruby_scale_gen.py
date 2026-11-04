# -*- coding: utf-8 -*-
"""Does a ruby field's width use the character scale (w:w) of its base and of
its reading?

blind-J ja/creative__9aa11078: «図書館» (HGS創英角ﾎﾟｯﾌﾟ体 b 28pt, w 66, spacing
1pt) under «としょかん» (14pt, w 66, spacing 1pt). Word sets «や» right after
the base; Oxi spread the base as if the reading were full width (5 x 14 = 70
over a 58pt base) and left a gap.

Faithful slice of that document (styles, theme, settings, sectPr): one
paragraph per arm «[ruby]や本», base / reading w:w = 66/66, 100/100, 66/100,
100/66, all with spacing 1pt, and 66/66 with spacing 0. Readout: x of «図»
(field start) and of «や».
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/creative/9aa1107806831b73.docx'
OUT = 'tests/fixtures/ruby_scale'; os.makedirs(OUT, exist_ok=True)
FONT = 'HGS創英角ﾎﾟｯﾌﾟ体'


def rpr(sz, w, sp):
    return (f'<w:rPr><w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/><w:b/>'
            + (f'<w:spacing w:val="{sp}"/>' if sp else '') + f'<w:w w:val="{w}"/><w:sz w:val="{sz}"/></w:rPr>')


def para(bw, rw, sp):
    ruby = ('<w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="26"/><w:hpsRaise w:val="50"/>'
            '<w:hpsBaseText w:val="56"/><w:lid w:val="ja-JP"/></w:rubyPr>'
            f'<w:rt><w:r>{rpr(28, rw, sp)}<w:t>としょかん</w:t></w:r></w:rt>'
            f'<w:rubyBase><w:r>{rpr(56, bw, sp)}<w:t>図書館</w:t></w:r></w:rubyBase></w:ruby>')
    return (f'<w:p><w:pPr><w:spacing w:line="0" w:lineRule="atLeast"/></w:pPr><w:r>{rpr(56, bw, sp)}{ruby}</w:r>'
            f'<w:r>{rpr(56, bw, sp)}<w:t>や本</w:t></w:r></w:p>')


ARMS = [(66, 66, 20), (100, 100, 20), (66, 100, 20), (100, 66, 20), (66, 66, 0), (66, 66, 40), (100, 100, 40), (100, 100, 0)]


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
                x = x[:b0] + ''.join(para(*a) for a in ARMS) + sect + x[x.index('</w:body>') + len('</w:body>'):]
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
    for b in P.get_text('rawdict')['blocks']:
        for l in b.get('lines', []):
            for s in l['spans']:
                for c in s['chars']:
                    if c['c'] in ('図', 'や'):
                        rows.append((round(c['origin'][1]), c['c'], round(c['origin'][0], 2)))
    by = {}
    for y, ch, x in sorted(rows):
        by.setdefault(y, {})[ch] = x
    for a, (y, v) in zip(ARMS, sorted(by.items())):
        print(f'base w{a[0]} rt w{a[1]} sp{a[2]}: 図 {v.get("図")}  や {v.get("や")}  field {v.get("や", 0) - v.get("図", 0):.2f}')
