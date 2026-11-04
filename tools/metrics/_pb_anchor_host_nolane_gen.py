# -*- coding: utf-8 -*-
"""Where does Word set the anchoring paragraph's own line when its square
floats leave it no lane?

blind-J en/creative__045bc787: the body is ONE paragraph (style Header, no text
of its own) anchoring Text Box 3 (324 x 646.5, margin-left, square), Text Box
2 (198 x 564.75 at column + 342, square) and Text Box 4; the page body is
108..756. Word keeps one page; Oxi sends the paragraph's line below the 646pt
box, past the page bottom, and the boxes follow it to page 2.

Arms on the document itself: as is / «ZZZ» text appended to the paragraph /
the paragraph followed by a second paragraph «YYY». Readout: page count and
where ZZZ / YYY land (page, y).
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/en/creative/045bc787d68b33a5.docx'
OUT = 'tests/fixtures/anchor_host_nolane'; os.makedirs(OUT, exist_ok=True)
ARMS = {'asis': None, 'text': 'ZZZ', 'next': 'YYY'}


def build(name, kind):
    zin = zipfile.ZipFile(SRC)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, data in items:
            if it.filename == 'word/document.xml' and kind:
                x = data.decode('utf8')
                i = x.index('<w:sectPr', x.index('<w:body>'))
                pend = x.rfind('</w:p>', 0, i)
                if kind == 'ZZZ':
                    x = x[:pend] + '<w:r><w:t>ZZZ</w:t></w:r>' + x[pend:]
                else:
                    x = x[:pend + 6] + '<w:p><w:r><w:t>YYY</w:t></w:r></w:p>' + x[pend + 6:]
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
    for n, k in ARMS.items():
        at = build(n, k)
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
        hits = []
        for pg, page in enumerate(P):
            for b in page.get_text('dict')['blocks']:
                for l in b.get('lines', []):
                    for s in l['spans']:
                        if s['text'].strip() in ('ZZZ', 'YYY'):
                            hits.append((pg + 1, round(s['origin'][0], 1), round(s['origin'][1], 1)))
        print(f'{n:5} pages {len(P)}  marker {hits}', flush=True)
