# -*- coding: utf-8 -*-
"""How much bottom room does Word reserve when a bordered table row splits at a page end?

en/educational__0079c94e (blind-J): the long table's row runs off p3; its line
«Illinois State University:» ends at 527.7 and Word keeps it on p3 (footer top
528.0) while Oxi splits at 527.2 (page bottom 527.7 - the S1188 0.5pt border
reserve) and carries it to p4, one page behind to the end.

The whole document, with only these changed per arm: the table borders (as
saved / tblBorders all none / all sz 24) and the footer (as saved / removed).
Readout: the page and baseline of «Illinois State University:».
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/en/educational/0079c94e815e3e76.docx'
OUT = 'tests/fixtures/split_border_reserve'; os.makedirs(OUT, exist_ok=True)
ARMS = {'orig': ('keep', True), 'bdr_none': ('none', True), 'bdr_24': ('24', True),
        'nofoot': ('keep', False)}
KEY = 'Illinois State University:'


def borders(x, mode):
    if mode == 'keep':
        return x
    def fix(m):
        s = m.group(0)
        if mode == 'none':
            return re.sub(r'w:val="[^"]*"', 'w:val="none"', s)
        return re.sub(r'w:sz="\d+"', 'w:sz="24"', s)
    return re.sub(r'<w:tblBorders>.*?</w:tblBorders>|<w:tcBorders>.*?</w:tcBorders>', fix, x, flags=re.S)


def build(name, mode, foot):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                x = borders(data.decode('utf8'), mode)
                if not foot:
                    x = re.sub(r'<w:footerReference [^>]*/>', '', x)
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    paths = {n: build(n, *a) for n, a in ARMS.items()}
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
        P = pymupdf.open(at[:-5] + '.pdf')
        hits = []
        for pi, page in enumerate(P):
            for b in page.get_text('dict')['blocks']:
                for l in b.get('lines', []):
                    t = ''.join(s['text'] for s in l['spans'])
                    if KEY in t:
                        hits.append((pi + 1, round(l['spans'][0]['origin'][1], 2), round(l['bbox'][3], 2)))
        print(f'{n:9} pages {len(P)}  {KEY!r} at {hits[:3]}', flush=True)
