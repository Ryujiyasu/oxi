# -*- coding: utf-8 -*-
"""How much will Word push in (追い込み) to keep a line's last character when the
line holds one «・» / one «、» / no mark?

blind-J ja/reference__6c5dcb56: «2020-2021年度　4R2Z　ゾーン・チェアパーソン»
(HG丸ｺﾞｼｯｸM-PRO 12, compressPunctuation, compat 15, no-type docGrid) fits in
Word with the «・» set at half width (6pt pushed in); Oxi's S595/S1487 tolerance
is min(4pt, marks x 0.5 em) and wraps the final «ン».

Faithful slice of that document (styles, settings, theme; one column):
paragraphs of 26 full-width characters (12pt = 312pt natural) with the mark at
position 12, right indent set so the line is short by D pt. Readout: does
character 25 stay on line 1.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/reference/6c5dcb5663cdd5fb.docx'
OUT = 'tests/fixtures/dot_oikomi_cap'; os.makedirs(OUT, exist_ok=True)
F = '<w:rFonts w:ascii="HG丸ｺﾞｼｯｸM-PRO" w:eastAsia="HG丸ｺﾞｼｯｸM-PRO" w:hAnsi="HG丸ｺﾞｼｯｸM-PRO"/><w:sz w:val="24"/>'
MARKS = {'dot': '・', 'comma': '、', 'none': 'ア'}
DS = [2, 3, 4, 5, 5.5, 6, 6.5, 7, 9, 12]
TEXT_W = 595.3 - 72 * 2  # page 11906tw with 1440tw margins


def para(mark, d):
    body = 'ゾーンチェアパーソンゾー' + mark + 'ンチェアパーソンゾーンチ' + 'Ｑ'
    right_tw = int(round((TEXT_W - (312 - d)) * 20))
    return (f'<w:p><w:pPr><w:ind w:left="0" w:right="{right_tw}"/><w:jc w:val="both"/><w:rPr>{F}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{F}</w:rPr><w:t>{body}</w:t></w:r></w:p>')


def build():
    zin = zipfile.ZipFile(SRC)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    at = os.path.abspath(os.path.join(OUT, 'arms.docx'))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
            'w:left="1440" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:linePitch="360"/></w:sectPr>')
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, data in items:
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                b0 = x.index('<w:body>') + len('<w:body>')
                body = ''.join(para(m, d) for m in MARKS.values() for d in DS)
                x = x[:b0] + body + sect + x[x.index('</w:body>'):]
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
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
    P = pymupdf.open(at[:-5] + '.pdf')
    lines = []
    for page in P:
        for b in page.get_text('dict')['blocks']:
            for l in b.get('lines', []):
                t = ''.join(s['text'] for s in l['spans']).strip()
                if t:
                    lines.append(t)
    # each paragraph is 1 line (kept) or 2 lines (wrapped: second line is the tail)
    i = 0
    for k, m in MARKS.items():
        for d in DS:
            first = lines[i]
            kept = first.endswith('Ｑ')
            i += 1 if kept else 2
            print(f'{k:6} short {d:4}: {"KEPT" if kept else "wrap"}  ({len(first)} chars)')
