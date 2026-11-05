# -*- coding: utf-8 -*-
"""Where does Word put a ruby line inside a text box?

ja/administrative__1fc69a4f (blind-J) p3: the «出張ういんぐ・関町» text box
(HG丸ｺﾞｼｯｸM-PRO 12pt, line 400 exact, snapToGrid 0, rt hps 12 raise 22):
Word's base baseline is 16.0 below the line top (the run border's box), Oxi's
28.0; the auto-spaced ruby paragraphs below advance 20.5 in Word, 15.6 in Oxi.

Faithful slice of that document: the box's content is replaced by «前の行» /
the ruby line «出張ういんぐ» / «後の行».  Arms <line>_<ruby>: line = ex400 /
auto, ruby = ruby / plain.  Readout: the baselines of 前, 出 (base), the ruby
«しゅっちょう» and 後, absolute on the page.

    python tools/metrics/_pb_txbx_ruby_gen.py            # build + Word PDF
    python tools/metrics/_pb_txbx_ruby_gen.py --oxi <renderer.exe>
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
SRC = 'pipeline_data/docx_corpus/ja/administrative/1fc69a4f23b5918f.docx'
OUT = 'tests/fixtures/txbx_ruby'; os.makedirs(OUT, exist_ok=True)
LINES = {'ex400': '<w:spacing w:line="400" w:lineRule="exact"/>', 'auto': ''}
ARMS = [f'{l}_{r}' for l in LINES for r in ('ruby', 'plain')]
F = 'HG丸ｺﾞｼｯｸM-PRO'
RPR = f'<w:rPr><w:rFonts w:ascii="{F}" w:eastAsia="{F}" w:hAnsi="{F}"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
RT = f'<w:rPr><w:rFonts w:ascii="{F}" w:eastAsia="{F}" w:hAnsi="{F}"/><w:sz w:val="12"/><w:szCs w:val="24"/></w:rPr>'


def content(line, ruby):
    ppr = f'<w:pPr><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/>{LINES[line]}{RPR}</w:pPr>'
    if ruby == 'ruby':
        mid = (f'<w:r>{RPR}<w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="12"/>'
               f'<w:hpsRaise w:val="22"/><w:hpsBaseText w:val="24"/><w:lid w:val="ja-JP"/></w:rubyPr>'
               f'<w:rt><w:r>{RT}<w:t>しゅっちょう</w:t></w:r></w:rt>'
               f'<w:rubyBase><w:r>{RPR}<w:t>出張</w:t></w:r></w:rubyBase></w:ruby></w:r>')
    else:
        mid = f'<w:r>{RPR}<w:t>出張</w:t></w:r>'
    return (f'<w:p>{ppr}<w:r>{RPR}<w:t>前の行</w:t></w:r></w:p>'
            f'<w:p>{ppr}{mid}<w:r>{RPR}<w:t>ういんぐ</w:t></w:r></w:p>'
            f'<w:p>{ppr}<w:r>{RPR}<w:t>後の行</w:t></w:r></w:p>')


def build(arm):
    line, ruby = arm.split('_')
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, arm + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                n = 0

                def repl(m):
                    nonlocal n
                    if '3929' in m.group(0):
                        n += 1
                        return '<w:txbxContent>' + content(line, ruby) + '</w:txbxContent>'
                    return m.group(0)
                x = re.sub(r'<w:txbxContent>.*?</w:txbxContent>', repl, x, flags=re.S)
                assert n >= 1, n
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


KEYS = ('前の行', '出張', 'しゅっちょう', '後の行')


def word_base(pdf):
    import pymupdf
    page = pymupdf.open(pdf)[2]
    base = {}
    for t in page.get_texttrace():
        s = ''.join(chr(c[0]) for c in t['chars'])
        for k in KEYS:
            if s.startswith(k) and t['chars'][0][2][0] < 290:
                base.setdefault(k, round(t['chars'][0][2][1], 2))
    return base


def oxi_base(exe, at):
    import json, subprocess, tempfile
    d = tempfile.mkdtemp()
    subprocess.run([exe, at, d + '/p', '72', '--dump-glyphs=' + d + '/g.json'], capture_output=True,
                   env=dict(os.environ, OXI_IGNORE_LRPB='1'))
    g = json.load(open(d + '/g.json', encoding='utf8'))
    pg = g['pages'][2] if isinstance(g, dict) else g[2]
    glyphs = pg['glyphs'] if isinstance(pg, dict) else pg
    base = {}
    for k in KEYS:
        for e in glyphs:
            if e.get('char', '') == k[0] and e['x'] < 290 and 360 < e.get('baseline', e.get('y', 0)) < 520:
                base.setdefault(k, round(e.get('baseline', e.get('y')), 2))
    return base


if __name__ == '__main__':
    paths = {a: build(a) for a in ARMS}
    if '--oxi' in sys.argv:
        exe = os.path.abspath(sys.argv[sys.argv.index('--oxi') + 1])
        for a, at in paths.items():
            w = word_base(at[:-5] + '.pdf') if os.path.exists(at[:-5] + '.pdf') else {}
            print(f'{a:12} oxi {oxi_base(exe, at)}  word {w}', flush=True)
        sys.exit()
    import win32com.client
    for a, at in paths.items():
        app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(at[:-5] + '.pdf', 17)
            finally:
                d.Close(False)
        finally:
            app.Quit()
        print(f'{a:12} word {word_base(at[:-5] + ".pdf")}', flush=True)
