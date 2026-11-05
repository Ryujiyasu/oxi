# -*- coding: utf-8 -*-
"""What advance does an ASCII space take in a w:hint="eastAsia" Arial run?

ja/reports__7a8268ce (blind-J p2-p6): the form rows «神経内科医»[fitText] then
«  » / «　»« » then «（». Word's PDF draws those spaces in ArialMT but advances
them 5.52 (= 11.04 / 2, the eastAsia half width); Oxi used Arial's 3.06, so
every «（人数人）» sat 2.5pt left (and the two-space row 3pt right).

Faithful slice of that document (styles, theme, settings, sectPr): one
paragraph per arm «医» + RUN + «（», RUN = Arial 11 text with/without the hint,
the document's balanceSingleByteDoubleByteWidth kept or removed.
Readout: «（» x - «医» x - the «医» advance (the a_sp_b / sp_a / a_sp arms read
the Latin-space context: a space between Latin letters keeps Arial's width in
the balance documents of the blind sets).
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/reports/7a8268ce7d484067.docx'
OUT = 'tests/fixtures/hint_ea_space'; os.makedirs(OUT, exist_ok=True)

RUNS = [('sp_hint', ' ', True), ('sp2_hint', '  ', True), ('sp_nohint', ' ', False),
        ('d3_hint', '3', True), ('d3_nohint', '3', False), ('ab_hint', 'ab', True),
        ('a_sp_b', 'a b', False), ('sp_a', ' a', False), ('a_sp', 'a ', False), ('ab_nohint', 'ab', False)]


def rpr(hint):
    h = ' w:hint="eastAsia"' if hint else ''
    return f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"{h}/><w:sz w:val="22"/></w:rPr>'


def para(text, hint):
    return ('<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr>'
            f'<w:r>{rpr(True)}<w:t>医</w:t></w:r>'
            f'<w:r>{rpr(hint)}<w:t xml:space="preserve">{text}</w:t></w:r>'
            f'<w:r>{rpr(True)}<w:t>（</w:t></w:r></w:p>')


def build(name, balance):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                b0 = x.index('<w:body>') + len('<w:body>')
                sect = re.search(r'<w:sectPr[ >].*?</w:sectPr>\s*</w:body>', x, re.S).group(0)
                body = ''.join(para(t, h) for _, t, h in RUNS)
                x = x[:b0] + body + sect + x[x.index('</w:body>') + len('</w:body>'):]
                data = x.encode('utf8')
            elif it.filename == 'word/settings.xml' and not balance:
                data = data.decode('utf8').replace('<w:balanceSingleByteDoubleByteWidth/>', '').encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    paths = {n: build(n, b) for n, b in (('balance', True), ('nobalance', False))}
    if '--build' in sys.argv:
        sys.exit()
    import win32com.client, pymupdf
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
        page = pymupdf.open(at[:-5] + '.pdf')[0]
        lines = {}
        for t in page.get_texttrace():
            for c in t['chars']:
                lines.setdefault(round(c[2][1]), []).append((c[2][0], chr(c[0]), t['font']))
        for (arm, _, _), y in zip(RUNS, sorted(lines)):
            cs = sorted(lines[y])
            x_i = next(x for x, ch, _ in cs if ch == '医')
            x_p = next(x for x, ch, _ in cs if ch == '（')
            fonts = [f for _, ch, f in cs if ch not in '医（']
            print(f'{n:10} {arm:10} run advance {x_p - x_i - 11.04:6.2f}  fonts {fonts}', flush=True)
