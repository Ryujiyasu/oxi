# -*- coding: utf-8 -*-
"""Does a table's last-row CELL bottom border set the room below the table?

blind-J ja/reports__7a8268ce p5: tables with tblBorders sz4 whose last row's
cells override the bottom with sz12 (1.5pt).  The «（注）» paragraph after
each table sits 1.0pt lower in Word than in Oxi, and the drift adds up per table.

Faithful slice; arms: the last-row cell bottoms sz 12 (as written) / 4 / 24 /
removed (tblBorders 0.5 applies).  Readout: the first table's last-row text
baseline -> «（注）» baseline.

    python tools/metrics/_pb_tbl_cellbottom_gen.py            # build + Word PDF
    python tools/metrics/_pb_tbl_cellbottom_gen.py --oxi <renderer.exe>
"""
import os, re, sys, zipfile, json, subprocess, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/reports/7a8268ce7d484067.docx'
OUT = 'tests/fixtures/tbl_cellbottom'; os.makedirs(OUT, exist_ok=True)
ARMS = ['sz12', 'sz4', 'sz24', 'none']


def build(arm):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, arm + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                i = x.find('貴施設における急性期脳卒中')
                t0 = x.rfind('<w:tbl>', 0, i); t1 = x.find('</w:tbl>', i)
                tbl = x[t0:t1]
                rows = re.findall(r'<w:tr[ >].*?</w:tr>', tbl, flags=re.S)
                last = rows[-1]
                if arm == 'none':
                    new = re.sub(r'<w:bottom w:val="single" w:sz="12"[^>]*/>', '', last)
                else:
                    new = last.replace('<w:bottom w:val="single" w:sz="12"', f'<w:bottom w:val="single" w:sz="{arm[2:]}"')
                x = x[:t0] + tbl.replace(last, new) + x[t1:]
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


def word_read(pdf):
    import pymupdf
    for pg in pymupdf.open(pdf):
        ys = {}
        for t in pg.get_texttrace():
            s = ''.join(chr(c[0]) for c in t['chars'])
            y = t['chars'][0][2][1]
            if '貴施設' in s: ys.setdefault('q', y)
            if s.startswith('（注）') and 'q' in ys: ys.setdefault('n', y)
        if 'q' in ys and 'n' in ys:
            return round(ys['n'] - ys['q'], 2)


def oxi_read(exe, at):
    d = tempfile.mkdtemp()
    subprocess.run([exe, at, d + '/p', '72', '--dump-glyphs=' + d + '/g.json'], capture_output=True,
                   env=dict(os.environ, OXI_IGNORE_LRPB='1'))
    for pg in json.load(open(d + '/g.json', encoding='utf8'))['pages']:
        g = pg['glyphs']
        txt = ''.join(e['char'] for e in g)
        if '貴施設' not in txt:
            continue
        i = txt.find('貴施設'); q = g[i]['baseline']
        j = txt.find('（注）', i)
        return round(g[j]['baseline'] - q, 2) if j >= 0 else None


if __name__ == '__main__':
    paths = {a: build(a) for a in ARMS}
    if '--oxi' in sys.argv:
        exe = os.path.abspath(sys.argv[sys.argv.index('--oxi') + 1])
        for a, at in paths.items():
            w = word_read(at[:-5] + '.pdf') if os.path.exists(at[:-5] + '.pdf') else None
            print(f'{a:6} oxi {oxi_read(exe, at)}  word {w}', flush=True)
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
        print(f'{a:6} word {word_read(at[:-5] + ".pdf")}', flush=True)
