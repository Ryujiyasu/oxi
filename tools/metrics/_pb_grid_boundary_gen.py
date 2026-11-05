# -*- coding: utf-8 -*-
"""Does a line whose 83/64 box is a hair over the grid pitch take two cells?

blind-J ja/educational__2aa00bc4 p2: an empty paragraph with a 12pt
«UD デジタル 教科書体 NP-R» mark (win 1802+567 / 2048 -> box 12 x 1.1567 x
83/64 = 18.0017) on a lines grid of 18: Word gives it 36 (two cells), Oxi 18,
so the left column held one more line in Oxi and «チャレンジ» stayed there.

Arms: <size>_<kind>: size = 11 / 12 / 13 (half-points 22 / 24 / 26), kind =
empty (mark only) / text («あいう»).  Each arm: «前» line (MS Mincho 10.5), the
arm paragraph, «後» line.  Readout: «前» -> «後» baseline gap (Word PDF / Oxi).

    python tools/metrics/_pb_grid_boundary_gen.py            # build + Word PDF
    python tools/metrics/_pb_grid_boundary_gen.py --oxi <renderer.exe>
"""
import os, sys, zipfile, json, subprocess, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

OUT = 'tests/fixtures/grid_boundary'; os.makedirs(OUT, exist_ok=True)
UD = 'UD デジタル 教科書体 NP-R'
ARMS = [f'{s}_{k}' for s in (22, 24, 26) for k in ('empty', 'text')] + [
    f'24_{k}_{c}_{f}' for k in ('empty', 'text') for c in ('c15', 'leg') for f in ('np', 'nk')]
FACES = {'np': 'UD デジタル 教科書体 NP-R', 'nk': 'UD デジタル 教科書体 NK-R'}


def rpr(font, sz):
    return f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz}"/></w:rPr>'


def doc(arm):
    sz, kind = arm.split('_')[:2]
    ud = FACES[arm.split('_')[3]] if arm.count('_') == 3 else UD
    m = 'ＭＳ 明朝'
    mid = f'<w:p><w:pPr>{rpr(ud, sz)}</w:pPr>' + (f'<w:r>{rpr(ud, sz)}<w:t>あいう</w:t></w:r>' if kind == 'text' else '') + '</w:p>'
    body = (f'<w:p><w:pPr>{rpr(m, 21)}</w:pPr><w:r>{rpr(m, 21)}<w:t>前の行</w:t></w:r></w:p>' + mid +
            f'<w:p><w:pPr>{rpr(m, 21)}</w:pPr><w:r>{rpr(m, 21)}<w:t>後の行</w:t></w:r></w:p>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440"'
            ' w:left="1080" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def build():
    ct = T.CT.replace('</Types>', '<Override PartName="/word/settings.xml" ContentType="application/vnd.'
                      'openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace('</Relationships>', '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                            'officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat>'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
                '</w:compat></w:settings>')
    legacy = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat>'
              '<w:balanceSingleByteDoubleByteWidth/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/>'
              '</w:compat></w:settings>')
    paths = {}
    for arm in ARMS:
        at = os.path.abspath(os.path.join(OUT, arm + '.docx'))
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', T.RELS)
            z.writestr('word/_rels/document.xml.rels', drels); z.writestr('word/styles.xml', T.STYLES)
            z.writestr('word/settings.xml', legacy if '_leg_' in arm else settings)
            z.writestr('word/document.xml', doc(arm))
        paths[arm] = at
    return paths


def word_gap(pdf):
    import pymupdf
    base = {}
    for t in pymupdf.open(pdf)[0].get_texttrace():
        s = ''.join(chr(c[0]) for c in t['chars'])
        for k in ('前の行', '後の行'):
            if s.startswith(k):
                base.setdefault(k, t['chars'][0][2][1])
    return round(base['後の行'] - base['前の行'], 2)


def oxi_gap(exe, at):
    d = tempfile.mkdtemp()
    subprocess.run([exe, at, d + '/p', '72', '--dump-glyphs=' + d + '/g.json'], capture_output=True)
    g = json.load(open(d + '/g.json', encoding='utf8'))['pages'][0]['glyphs']
    y = {e['char']: e['baseline'] for e in g if e['char'] in '前後'}
    return round(y['後'] - y['前'], 2)


if __name__ == '__main__':
    paths = build()
    if '--oxi' in sys.argv:
        exe = os.path.abspath(sys.argv[sys.argv.index('--oxi') + 1])
        for a, at in paths.items():
            w = word_gap(at[:-5] + '.pdf') if os.path.exists(at[:-5] + '.pdf') else None
            print(f'{a:10} oxi {oxi_gap(exe, at)}  word {w}', flush=True)
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
        print(f'{a:10} word {word_gap(at[:-5] + ".pdf")}', flush=True)
