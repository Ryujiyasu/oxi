# -*- coding: utf-8 -*-
"""Where does Word wrap «2020-2021年度　4R2Z　ゾーン・チェアパーソン» as its line
narrows? The line mixes Latin and kana (two autoSpace gaps) and one «・».

blind-J ja/reference__6c5dcb56: Word keeps the final «ン» on a 288.6pt line
(the «・» pushed in to half width and the 1|年 gap at 2.27 instead of 3); Oxi
wraps it. Faithful slice (styles, settings, theme; one column, jc both): the
real paragraph text (runs split as in the source) with the line width W swept
in 0.5pt steps. Readout: the first W at which «ン» stays on line 1.
Arms: as written / «・» replaced by «ア» / Latin replaced by full-width digits.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/reference/6c5dcb5663cdd5fb.docx'
OUT = 'tests/fixtures/latin_dot_line'; os.makedirs(OUT, exist_ok=True)
F = ('<w:rFonts w:ascii="HG丸ｺﾞｼｯｸM-PRO" w:eastAsia="HG丸ｺﾞｼｯｸM-PRO" w:hAnsi="HG丸ｺﾞｼｯｸM-PRO" w:hint="eastAsia"/>'
     '<w:sz w:val="24"/><w:szCs w:val="24"/>')
TEXTS = {
    'real': ['2020', '-2', '021', '年度　', '4R2Z　ゾーン・チェアパーソン'],
    'nodot': ['2020', '-2', '021', '年度　', '4R2Z　ゾーンアチェアパーソン'],
}
WS = [w / 2 for w in range(2 * 280, 2 * 300 + 1)]
TEXT_W = 595.3 - 72 * 2


def para(runs, w):
    right_tw = int(round((TEXT_W - w) * 20))
    rr = ''.join(f'<w:r><w:rPr>{F}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>' for t in runs)
    return f'<w:p><w:pPr><w:ind w:left="0" w:right="{right_tw}"/><w:jc w:val="both"/><w:rPr>{F}</w:rPr></w:pPr>{rr}</w:p>'


def build():
    zin = zipfile.ZipFile(SRC)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    at = os.path.abspath(os.path.join(OUT, 'arms.docx'))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" '
            'w:left="1440" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:linePitch="360"/></w:sectPr>')
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, data in items:
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                b0 = x.index('<w:body>') + len('<w:body>')
                body = ''.join(para(r, w) for r in TEXTS.values() for w in WS)
                x = x[:b0] + body + sect + x[x.index('</w:body>'):]
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


def first_lines(pdf):
    import pymupdf
    out = []
    for page in pymupdf.open(pdf):
        for b in page.get_text('dict')['blocks']:
            for l in b.get('lines', []):
                t = ''.join(s['text'] for s in l['spans']).strip()
                if t:
                    out.append(t)
    return out


if __name__ == '__main__':
    import win32com.client
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
    lines = first_lines(at[:-5] + '.pdf')
    i = 0
    for k in TEXTS:
        kept = []
        for w in WS:
            ok = lines[i].endswith('ソン')
            kept.append((w, ok))
            i += 1 if ok else 2
        first_ok = next((w for w, ok in kept if ok), None)
        print(f'{k:6} first width that keeps «ン»: {first_ok}')
