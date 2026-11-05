# -*- coding: utf-8 -*-
"""Where does Word stack a page's footnote block?

ja/reports__75ebb654 (blind-J p1): Word's separator sits at 654.7 and the
footnote's first line at 662.6 (18.6 tall, the next ones 17.3); Oxi drew the
block 5.8pt lower. The document: A4, margins 1134tw, docGrid lines 360, the
separator paragraph before=120, footnote text style «Kaiti TC» 10 (missing,
Word shows Microsoft JhengHei), footnote = one text paragraph + one empty one.

Faithful slice of that document (styles, theme, settings, fontTable, sectPr,
footnotes part rewritten per arm): body = «本文» + footnote reference.
Arms: the footnote's paragraphs (text lines x empty paragraph) and whether
the reference mark is superscripted. Footer arms (on 'one'): no footerReference,
footer paragraphs before 0 / 24. Readout: separator rule y, each footnote
line's top, the last line's bottom.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/reports/75ebb6540d49b937.docx'
OUT = 'tests/fixtures/footnote_stack'; os.makedirs(OUT, exist_ok=True)
LINE = '註腳文字測試註腳文字測試註腳文字測試註腳文字測試註腳文字測試註腳文字測試註腳文字測試'

# name -> (text paragraphs, trailing empty paragraph, reference rStyle)
ARMS = {
    'one': (1, False, True),
    'one_empty': (1, True, True),
    'two': (2, False, True),
    'one_plainref': (1, False, False),
    'orig_shape': (1, True, True),
}
# footer variants run on the 'one' footnote: keep / no footerReference / footer before 0
FOOTERS = {'nofoot': 'none', 'foot_b0': 'b0', 'foot_b24': 'b24'}


def footnotes_xml(x, n_text, empty, styled):
    head = x[:x.index('<w:footnote w:id="1"')]
    ref = '<w:r><w:rPr><w:rStyle w:val="af4"/></w:rPr><w:footnoteRef/></w:r>' if styled \
        else '<w:r><w:footnoteRef/></w:r>'
    paras = []
    for k in range(n_text):
        r = ref + '<w:r><w:t xml:space="preserve"> </w:t></w:r>' if k == 0 else ''
        paras.append(f'<w:p><w:pPr><w:pStyle w:val="af2"/></w:pPr>{r}<w:r><w:t>{LINE[:30]}</w:t></w:r></w:p>')
    if empty:
        paras.append('<w:p><w:pPr><w:pStyle w:val="af2"/></w:pPr></w:p>')
    return head + '<w:footnote w:id="1">' + ''.join(paras) + '</w:footnote></w:footnotes>'


def build(name, arm, footer='keep'):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                b0 = x.index('<w:body>') + len('<w:body>')
                sect = re.search(r'<w:sectPr[ >].*?</w:sectPr>\s*</w:body>', x, re.S).group(0)
                body = ('<w:p><w:r><w:t>本文</w:t></w:r><w:r><w:rPr><w:rStyle w:val="af4"/></w:rPr>'
                        '<w:footnoteReference w:id="1"/></w:r></w:p>')
                if footer == 'none':
                    sect = re.sub(r'<w:footerReference [^>]*/>', '', sect)
                x = x[:b0] + body + sect + x[x.index('</w:body>') + len('</w:body>'):]
                data = x.encode('utf8')
            elif it.filename.startswith('word/footer') and footer in ('b0', 'b24'):
                v = {'b0': '0', 'b24': '480'}[footer]
                data = data.decode('utf8').replace('<w:spacing w:before="120"/>',
                    f'<w:spacing w:before="{v}" w:beforeLines="0"/>').encode('utf8')
            elif it.filename == 'word/footnotes.xml':
                data = footnotes_xml(data.decode('utf8'), *arm).encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    paths = {n: build(n, a) for n, a in ARMS.items()}
    paths.update({n: build(n, ARMS['one'], f) for n, f in FOOTERS.items()})
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
        rule = [round(dr['rect'].y0, 2) for dr in page.get_drawings() if dr['rect'].y0 > 400]
        lines = [(round(l['bbox'][1], 2), round(l['bbox'][3], 2), ''.join(s['text'] for s in l['spans'])[:6])
                 for b in page.get_text('dict')['blocks'] for l in b.get('lines', []) if l['bbox'][1] > 400]
        print(f'{n:13} rule {rule}  lines {lines}', flush=True)
