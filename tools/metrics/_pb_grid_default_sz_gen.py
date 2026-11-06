# -*- coding: utf-8 -*-
"""Which default size sets a character grid's pitch when no style names one?

blind-J ja/legal__10f9c779: docGrid linesAndChars with no charSpace, a
docDefaults rPrDefault without w:sz and a Normal style without w:sz.  Word
sets the body's 10pt kanji on a 10.0pt pitch; Oxi's parser took 10.5 as the
grid's default size, so every line carried 0.5pt per char of grid extra and
compressed its 、 to fit.

Arms <where the default size is declared>_<charSpace>: none (no sz anywhere),
dd21 / dd20 (rPrDefault sz), n21 / n24 (Normal sz only).  The body run is an
explicit 10pt MS Mincho run.  Readout: the median kanji advance on line 1
(Word PDF / Oxi glyph dump).

    python tools/metrics/_pb_grid_default_sz_gen.py            # build + Word PDF
    python tools/metrics/_pb_grid_default_sz_gen.py --oxi <renderer.exe>
"""
import os, sys, zipfile, json, subprocess, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

OUT = 'tests/fixtures/grid_default_sz' + ('' if os.environ.get('PB_JC', 'left') == 'left' else '_' + os.environ['PB_JC']); os.makedirs(OUT, exist_ok=True)
ARMS = {f'{w}_{cs}': (w, cs) for w in ('none', 'dd21', 'dd20', 'n21', 'n24') for cs in (0, 2000)}
M = 'ＭＳ 明朝'
TEXT = '乙の事業範囲は前記のとおりであり乙の事業範囲を証するものとして許可証の写しを添付する' * 2


JC = os.environ.get("PB_JC", "left")


def styles(where):
    dsz = {'dd21': '<w:sz w:val="21"/>', 'dd20': '<w:sz w:val="20"/>'}.get(where, '')
    nsz = {'n21': '<w:sz w:val="21"/>', 'n24': '<w:sz w:val="24"/>'}.get(where, '')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults>'
            f'<w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="{M}" w:hAnsi="Century"/>{dsz}'
            '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            f'<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:jc w:val="{JC}"/>'
            f'</w:pPr><w:rPr>{nsz}</w:rPr></w:style></w:styles>')


def doc(cs):
    grid = f'<w:docGrid w:type="linesAndChars" w:linePitch="360"{f" w:charSpace={chr(34)}{cs}{chr(34)}" if cs else ""}/>'
    rpr = f'<w:rPr><w:rFonts w:ascii="{M}" w:eastAsia="{M}" w:hAnsi="{M}"/><w:sz w:val="20"/></w:rPr>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>'
            f'<w:p><w:r>{rpr}<w:t>{TEXT}</w:t></w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1248" w:bottom="1440"'
            ' w:left="1248" w:header="851" w:footer="992" w:gutter="0"/>' + grid + '</w:sectPr></w:body></w:document>')


def build():
    paths = {}
    for arm, (where, cs) in ARMS.items():
        at = os.path.abspath(os.path.join(OUT, arm + '.docx'))
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', T.CT); z.writestr('_rels/.rels', T.RELS)
            z.writestr('word/_rels/document.xml.rels', T.DRELS); z.writestr('word/styles.xml', styles(where))
            z.writestr('word/document.xml', doc(cs))
        paths[arm] = at
    return paths


def med(xs):
    adv = sorted(b - a for a, b in zip(xs, xs[1:]))
    return round(adv[len(adv) // 2], 2), len(xs)


def word_read(pdf):
    import pymupdf
    ch = [(c[2][1], c[2][0]) for t in pymupdf.open(pdf)[0].get_texttrace() for c in t['chars']]
    y0 = min(y for y, _ in ch)
    return med(sorted(x for y, x in ch if abs(y - y0) < 0.5))


def oxi_read(exe, at):
    d = tempfile.mkdtemp()
    subprocess.run([exe, at, d + '/p', '72', '--dump-glyphs=' + d + '/g.json'], capture_output=True)
    g = json.load(open(d + '/g.json', encoding='utf8'))['pages'][0]['glyphs']
    y0 = min(e['baseline'] for e in g)
    return med(sorted(e['x'] for e in g if abs(e['baseline'] - y0) < 0.5))


if __name__ == '__main__':
    paths = build()
    if '--oxi' in sys.argv:
        exe = os.path.abspath(sys.argv[sys.argv.index('--oxi') + 1])
        for n, at in paths.items():
            w = word_read(at[:-5] + '.pdf') if os.path.exists(at[:-5] + '.pdf') else None
            print(f'{n:10} oxi {oxi_read(exe, at)}  word {w}', flush=True)
        sys.exit()
    import win32com.client
    for n, at in paths.items():
        app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(at[:-5] + '.pdf', 17)
            finally:
                d.Close(False)
        finally:
            app.Quit()
        print(f'{n:10} word {word_read(at[:-5] + ".pdf")}', flush=True)
