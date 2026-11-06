# -*- coding: utf-8 -*-
"""How far below a text box's inline picture does the next paragraph start?

blind-J en/creative__045bc787: the left text box (compatLnSpc) holds an
image-only paragraph (Header style, 309.5 x 196.5 picture, docDefaults before /
after 200, line 276) and then «Cosmetology» (Arial 36 bold, before 200).
Word: picture bottom -> heading baseline 45.8; Oxi 53.8 (it adds the picture
paragraph's after and the heading's before, 10 + 10).

Faithful slice; the box content becomes the original picture paragraph (after
a) + «CCCC» (Arial 12, before b, line 240) + «DDDD».  Arms (a, b) in twips.
Readout: picture bottom -> «CCCC» baseline, «CCCC» -> «DDDD».

    python tools/metrics/_pb_txbx_img_gap_gen.py            # build + Word PDF
    python tools/metrics/_pb_txbx_img_gap_gen.py --oxi <renderer.exe>
"""
import os, re, sys, zipfile, json, subprocess, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/en/creative/045bc787d68b33a5.docx'
OUT = 'tests/fixtures/txbx_img_gap'; os.makedirs(OUT, exist_ok=True)
ARMS = {'a200_b200': (200, 200), 'a200_b0': (200, 0), 'a0_b200': (0, 200), 'a0_b0': (0, 0), 'a400_b200': (400, 200),
        'a200_b400': (200, 400), 'a0_b0_sz40': (0, 0), 'a0_b0_mark8': (0, 0)}


def para(text, before):
    return (f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>{text}</w:t></w:r></w:p>')


def build(name, a, b):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                n = 0

                def repl(m):
                    nonlocal n
                    if '>Cosmetology<' not in m.group(0):
                        return m.group(0)
                    n += 1
                    pic = re.search(r'<w:p [^>]*>(?:(?!</w:p>).)*?3930650.*?</w:p>', m.group(0), flags=re.S).group(0)
                    pic = pic.replace('<w:pStyle w:val="Header"/>',
                                      f'<w:pStyle w:val="Header"/><w:spacing w:after="{a}"/>', 1)
                    if name.endswith('_sz40'):
                        pic = pic.replace('<w:sz w:val="20"/>', '<w:sz w:val="40"/>', 1)
                    if name.endswith('_mark8'):
                        pic = pic.replace('<w:pStyle w:val="Header"/>', '<w:pStyle w:val="Header"/><w:rPr><w:sz w:val="16"/></w:rPr>', 1) if '<w:rPr>' not in pic.split('</w:pPr>')[0] else pic.replace('</w:rPr></w:pPr>', '<w:sz w:val="16"/></w:rPr></w:pPr>', 1)
                    return '<w:txbxContent>' + pic + para('CCCC', b) + para('DDDD', 0) + '</w:txbxContent>'
                x = re.sub(r'<w:txbxContent>.*?</w:txbxContent>', repl, x, flags=re.S)
                assert n >= 1, n
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


def word_read(pdf):
    import pymupdf
    pg = pymupdf.open(pdf)[0]
    img = [im['bbox'] for im in pg.get_image_info() if abs(im['bbox'][3] - im['bbox'][1] - 196.5) < 1.0]
    base = {}
    for t in pg.get_texttrace():
        s = ''.join(chr(c[0]) for c in t['chars'])
        for k in ('CCCC', 'DDDD'):
            if k in s:
                base.setdefault(k, t['chars'][0][2][1])
    if not img or len(base) < 2:
        return None
    return round(base['CCCC'] - img[0][3], 2), round(base['DDDD'] - base['CCCC'], 2)


def oxi_read(exe, at):
    d = tempfile.mkdtemp()
    for flag in ('--dump-glyphs=' + d + '/g.json', '--dump-layout=' + d + '/l.json'):
        subprocess.run([exe, at, d + '/p', '72', flag], capture_output=True, env=dict(os.environ, OXI_IGNORE_LRPB='1'))
    lay = json.load(open(d + '/l.json', encoding='utf8'))['pages'][0]['elements']
    img = [e for e in lay if e['type'] == 'image' and abs(e['h'] - 196.5) < 1.0]
    g = json.load(open(d + '/g.json', encoding='utf8'))['pages'][0]['glyphs']
    base = {}
    bottom = img[0]['y'] + img[0]['h'] if img else 0.0
    for e in g:
        if e['char'] in 'CD' and e['baseline'] > bottom and e['x'] < img[0]['x'] + img[0]['w']:
            base.setdefault(e['char'], e['baseline'])
    if not img or len(base) < 2:
        return None
    return round(base['C'] - img[0]['y'] - img[0]['h'], 2), round(base['D'] - base['C'], 2)


if __name__ == '__main__':
    paths = {n: build(n, *ab) for n, ab in ARMS.items()}
    if '--oxi' in sys.argv:
        exe = os.path.abspath(sys.argv[sys.argv.index('--oxi') + 1])
        for n, at in paths.items():
            w = word_read(at[:-5] + '.pdf') if os.path.exists(at[:-5] + '.pdf') else None
            print(f'{n:10} oxi {oxi_read(exe, at)}  word {w}', flush=True)
        sys.exit()
    import win32com.client
    only = [a for a in sys.argv[1:] if a in ARMS]
    for n, at in paths.items():
        if only and n not in only:
            continue
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
