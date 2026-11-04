# -*- coding: utf-8 -*-
"""Does an inline chart grow an EXACT body line?

ja/reports__7a8daf36 (blind-J): body paragraph 13 is `line=240 exact` and holds
text plus an inline chart 324 x 220.5pt. Word keeps the page to one page; Oxi
grew the line by the chart (next paragraph +221pt) and spilled onto page 2.
The S1320/S1321 law (an inline object on an exact line does not grow it) is
implemented for picture-only paragraphs and cells, not for an object run
beside text (the S851 growth path).

Faithful slice of that document (its styles, theme, settings, sectPr, chart
parts): body = «ああああ» / X / «いいいい», X = paragraph 13 itself with its
anchors removed, under its own exact 12pt / auto / exact 30pt.
Readout: the baseline gap «ああああ» -> «いいいい».
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/reports/7a8daf3659417f2a.docx'
OUT = 'tests/fixtures/exact_inline_chart'; os.makedirs(OUT, exist_ok=True)


def para13(x):
    from lxml import etree
    t = etree.fromstring(x.encode('utf8'))
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
          'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
          'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'}
    p = t.xpath('//wp:inline', namespaces=ns)[0].xpath('ancestor::w:p[1]', namespaces=ns)[0]
    for a in p.xpath('.//wp:anchor', namespaces=ns):
        r = a.xpath('ancestor::w:r[1]', namespaces=ns)[0]
        r.getparent().remove(r)
    for a in p.xpath('.//mc:AlternateContent', namespaces=ns):
        a.getparent().remove(a)
    return etree.tostring(p, encoding='unicode')


def text_p(t):
    return (f'<w:p><w:pPr><w:spacing w:after="0" w:line="0" w:lineRule="atLeast"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr><w:t>{t}</w:t></w:r></w:p>')


ARMS = {
    'exact12': None,
    'auto': '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
    'exact30': '<w:spacing w:after="0" w:line="600" w:lineRule="exact"/>',
    'exact12_noafter': '<w:spacing w:after="0" w:line="240" w:lineRule="exact"/>',
}


def build(name, spacing):
    zin = zipfile.ZipFile(SRC)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, data in items:
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                p = para13(x)
                if spacing:
                    p = re.sub(r'<w:spacing [^>]*/>', spacing, p, count=1)
                b0 = x.index('<w:body>') + len('<w:body>')
                sect = re.search(r'<w:sectPr[ >].*?</w:sectPr>\s*</w:body>', x, re.S).group(0)
                x = x[:b0] + text_p('ああああ') + p + text_p('いいいい') + sect + x[x.index('</w:body>') + len('</w:body>'):]
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
    paths = {n: build(n, s) for n, s in ARMS.items()}
    if '--build' in sys.argv:
        sys.exit()
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
        P = pymupdf.open(at[:-5] + '.pdf')
        ys = {}
        for pg, page in enumerate(P):
            for b in page.get_text('dict')['blocks']:
                for l in b.get('lines', []):
                    for s in l['spans']:
                        t = s['text'].strip()
                        if t in ('ああああ', 'いいいい'):
                            ys[t] = (pg, s['origin'][1])
        print(f'{n:16} pages {len(P)}  A {ys.get("ああああ")}  B {ys.get("いいいい")}  gap {ys["いいいい"][1] - ys["ああああ"][1]:.2f}', flush=True)
