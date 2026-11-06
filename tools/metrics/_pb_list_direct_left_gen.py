# -*- coding: utf-8 -*-
"""Where does a bullet's text start when the paragraph overrides the level's left?

blind-J en/policies__00cb6db7 p6: bullets numId 13 (level ind left 720 hanging
360, num tab 720) with a direct <w:ind w:left="-225"/> (no hanging): Word puts
the bullet at margin - 29.25 and the text at margin - 11.25 (the paragraph's
left); Oxi (S778, S853's negative-marker arm) sent the text to the level's 36pt.

Faithful slice; the first bullet's direct ind is swept.  Arms: the direct
w:ind attributes.  Readout: the bullet x and the first word's x (Word PDF /
Oxi), relative to the left margin (72).

    python tools/metrics/_pb_list_direct_left_gen.py            # build + Word PDF
    python tools/metrics/_pb_list_direct_left_gen.py --oxi <renderer.exe>
"""
import os, re, sys, zipfile, json, subprocess, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/en/policies/00cb6db7c29b1ebc.docx'
OUT = 'tests/fixtures/list_direct_left'; os.makedirs(OUT, exist_ok=True)
ARMS = {
    'l-225': 'w:left="-225"',
    'l-225_h360': 'w:left="-225" w:hanging="360"',
    'l-225_h431': 'w:left="-225" w:hanging="431"',
    'l153_h431': 'w:left="153" w:hanging="431"',
    'l153': 'w:left="153"',
    'l300': 'w:left="300"',
    'l300_h360': 'w:left="300" w:hanging="360"',
    'l-225_f0': 'w:left="-225" w:firstLine="0"',
}
KEY = 'To reduce the amount'


def build(name, ind):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                i = x.find(KEY)
                a = x.rfind('<w:p ', 0, i)
                p0 = x[a:i]
                p1 = p0.replace('<w:ind w:left="-225"/>', f'<w:ind {ind}/>', 1)
                assert p1 != p0 or name == 'l-225'
                x = x[:a] + p1 + x[i:]
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


def word_read(pdf):
    import pymupdf
    for pg in pymupdf.open(pdf):
        tr = pg.get_texttrace()
        for k, t in enumerate(tr):
            s = ''.join(chr(c[0]) for c in t['chars'])
            if s.startswith('To reduce'):
                y = t['chars'][0][2][1]
                bul = [u['chars'][0][2][0] for u in tr if abs(u['chars'][0][2][1] - y) < 0.5
                       and chr(u['chars'][0][0]) in '•']
                return (round(bul[0] - 72, 2) if bul else None, round(t['chars'][0][2][0] - 72, 2))


def oxi_read(exe, at):
    d = tempfile.mkdtemp()
    subprocess.run([exe, at, d + '/p', '72', '--dump-layout=' + d + '/l.json'], capture_output=True,
                   env=dict(os.environ, OXI_IGNORE_LRPB='1'))
    for pg in json.load(open(d + '/l.json', encoding='utf8'))['pages']:
        els = [e for e in pg['elements'] if e['type'] == 'text']
        for k, e in enumerate(els):
            if e['text'] == 'To' and k + 2 < len(els) and els[k + 2]['text'] == 'reduce':
                bul = [b for b in els if abs(b['y'] - e['y']) < 0.5 and b['text'] in ('', '•')]
                return (round(bul[0]['x'] - 72, 2) if bul else None, round(e['x'] - 72, 2))


if __name__ == '__main__':
    paths = {n: build(n, ind) for n, ind in ARMS.items()}
    if '--oxi' in sys.argv:
        exe = os.path.abspath(sys.argv[sys.argv.index('--oxi') + 1])
        for n, at in paths.items():
            w = word_read(at[:-5] + '.pdf') if os.path.exists(at[:-5] + '.pdf') else None
            print(f'{n:12} oxi {oxi_read(exe, at)}  word {w}', flush=True)
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
        print(f'{n:12} word {word_read(at[:-5] + ".pdf")}', flush=True)
