# -*- coding: utf-8 -*-
"""Does an empty paragraph measure its ASCII face when fontTable declares that
face with a Japanese charset?

blind-J ja/correspondence__199abbcc: docDefaults ascii «Liberation Serif»
(installed; fontTable charset 80, altName Times New Roman) / eastAsia ＭＳ 明朝,
Normal 12pt, no-type grid 240.  The empty paragraphs between the numbered items
advance ~15.7 in Word (the ＭＳ 明朝 box 15.56), Oxi ~13.9 (the Liberation
Serif hhea box, S583 / S707's Western-ASCII empty mark).

Faithful slice; arms: the Liberation Serif fontTable charset 80 (as written) /
00 / the declaration removed.  Readout: the «２» -> «３» baseline gap (one
empty paragraph between them) and «３» -> «４» (none).  Each arm opens in a
fresh Word process.

    python tools/metrics/_pb_empty_ascii_charset_gen.py            # build + Word PDF
    python tools/metrics/_pb_empty_ascii_charset_gen.py --oxi <renderer.exe>
"""
import os, re, sys, zipfile, json, subprocess, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/correspondence/199abbcce9e5d851.docx'
OUT = 'tests/fixtures/empty_ascii_charset'; os.makedirs(OUT, exist_ok=True)
ARMS = ['cs80', 'cs00', 'nodecl', 'cs80_roman_fixed', 'cs80_swiss_variable', 'cs80_swiss_fixed']
KEYS = ('２　主', '３　後', '４　期')


def build(arm):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, arm + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/fontTable.xml':
                x = data.decode('utf8')
                m = re.search(r'<w:font w:name="Liberation Serif">.*?</w:font>', x, flags=re.S)
                assert m
                if arm == 'cs00':
                    x = x.replace(m.group(0), m.group(0).replace('<w:charset w:val="80"/>', '<w:charset w:val="00"/>'))
                elif arm == 'nodecl':
                    x = x.replace(m.group(0), '')
                elif arm.startswith('cs80_'):
                    fam, pitch = arm.split('_')[1:]
                    x = x.replace(m.group(0), m.group(0).replace('<w:family w:val="roman"/>', f'<w:family w:val="{fam}"/>')
                                  .replace('<w:pitch w:val="variable"/>', f'<w:pitch w:val="{pitch}"/>'))
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


def word_gaps(pdf):
    import pymupdf
    ys = {}
    for page in pymupdf.open(pdf):
        for t in page.get_texttrace():
            s = ''.join(chr(c[0]) for c in t['chars'])
            for k in KEYS:
                if s.startswith(k[0]) and k not in ys:
                    nxt = page.get_textbox(pymupdf.Rect(t['bbox'][0], t['bbox'][1], t['bbox'][0] + 60, t['bbox'][3]))
                    if k[2] in nxt:
                        ys[k] = t['chars'][0][2][1]
        if len(ys) == len(KEYS):
            break
        ys = {}
    v = [ys.get(k) for k in KEYS]
    return [round(v[1] - v[0], 2) if v[0] and v[1] else None, round(v[2] - v[1], 2) if v[1] and v[2] else None]


def oxi_gaps(exe, at):
    d = tempfile.mkdtemp()
    subprocess.run([exe, at, d + '/p', '72', '--dump-glyphs=' + d + '/g.json'], capture_output=True,
                   env=dict(os.environ, OXI_IGNORE_LRPB='1'))
    ys = {}
    for pg in json.load(open(d + '/g.json', encoding='utf8'))['pages']:
        g = pg['glyphs']
        ys = {}
        for i, e in enumerate(g):
            for k in KEYS:
                if e['char'] == k[0] and k not in ys and any(x['char'] == k[2] and abs(x['baseline'] - e['baseline']) < 0.5 for x in g[i:i + 4]):
                    ys[k] = e['baseline']
        if len(ys) == len(KEYS):
            break
    v = [ys.get(k) for k in KEYS]
    return [round(v[1] - v[0], 2) if v[0] and v[1] else None, round(v[2] - v[1], 2) if v[1] and v[2] else None]


if __name__ == '__main__':
    paths = {a: build(a) for a in ARMS}
    if '--oxi' in sys.argv:
        exe = os.path.abspath(sys.argv[sys.argv.index('--oxi') + 1])
        for a, at in paths.items():
            w = word_gaps(at[:-5] + '.pdf') if os.path.exists(at[:-5] + '.pdf') else None
            print(f'{a:8} oxi {oxi_gaps(exe, at)}  word {w}', flush=True)
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
        import pymupdf
        fonts = sorted({sp['font'] for pg in pymupdf.open(at[:-5] + '.pdf') for b in pg.get_text('dict')['blocks']
                        for l in b.get('lines', []) for sp in l['spans'] if 'E-mail' in sp['text'] or '0003' in sp['text']})
        print(f'{a:8} word {word_gaps(at[:-5] + ".pdf")} latin {fonts}', flush=True)
