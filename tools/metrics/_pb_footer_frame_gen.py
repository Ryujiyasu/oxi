# -*- coding: utf-8 -*-
"""Where does Word put a footer's framed page number (framePr vAnchor=text)?

blind-J ja/reports__75ebb654: footer2 = a framed «- 1 -» PAGE field
(framePr wrap=none vAnchor=text hAnchor=margin xAlign=center y=1, spacing
before 120) followed by an empty paragraph (before 120), style «ac»
(Microsoft JhengHei 12).  Word's baseline is 787.06, Oxi's 794.9.

Faithful slice of that document; the footer's two spacing-befores are swept.
Arms f<frame before>_e<empty before> in twips; one arm drops the empty
paragraph.  Readout: the «1» baseline on page 1 (Word PDF / Oxi glyph dump).

    python tools/metrics/_pb_footer_frame_gen.py            # build + Word PDF
    python tools/metrics/_pb_footer_frame_gen.py --oxi <renderer.exe>
"""
import os, re, sys, zipfile, json, subprocess, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/reports/75ebb6540d49b937.docx'
OUT = 'tests/fixtures/footer_frame'; os.makedirs(OUT, exist_ok=True)
ARMS = {'f120_e120': (120, 120), 'f0_e120': (0, 120), 'f120_e0': (120, 0), 'f0_e0': (0, 0),
        'f120_none': (120, None), 'f0_none': (0, None)}


def build(name, fb, eb):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/footer2.xml':
                x = data.decode('utf8')
                paras = re.findall(r'<w:p [^>]*>.*?</w:p>', x, flags=re.S)
                assert len(paras) == 2, len(paras)
                fr = paras[0].replace('<w:spacing w:before="120"/>', f'<w:spacing w:before="{fb}"/>')
                em = '' if eb is None else paras[1].replace('<w:spacing w:before="120"/>', f'<w:spacing w:before="{eb}"/>')
                x = x.replace(paras[0], fr).replace(paras[1], em)
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


def word_base(pdf):
    import pymupdf
    for t in pymupdf.open(pdf)[0].get_texttrace():
        s = ''.join(chr(c[0]) for c in t['chars'])
        if s.strip().startswith('1') and t['chars'][0][2][1] > 700:
            return round(t['chars'][0][2][1], 2)


def oxi_base(exe, at):
    d = tempfile.mkdtemp()
    subprocess.run([exe, at, d + '/p', '72', '--dump-glyphs=' + d + '/g.json'], capture_output=True,
                   env=dict(os.environ, OXI_IGNORE_LRPB='1'))
    g = json.load(open(d + '/g.json', encoding='utf8'))['pages'][0]['glyphs']
    ys = [e['baseline'] for e in g if e['char'] == '1' and e['baseline'] > 700]
    return round(ys[0], 2) if ys else None


if __name__ == '__main__':
    paths = {n: build(n, *v) for n, v in ARMS.items()}
    if '--oxi' in sys.argv:
        exe = os.path.abspath(sys.argv[sys.argv.index('--oxi') + 1])
        for n, at in paths.items():
            w = word_base(at[:-5] + '.pdf') if os.path.exists(at[:-5] + '.pdf') else None
            print(f'{n:10} oxi {oxi_base(exe, at)}  word {w}', flush=True)
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
        print(f'{n:10} word {word_base(at[:-5] + ".pdf")}', flush=True)
