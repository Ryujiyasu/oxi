# -*- coding: utf-8 -*-
"""Faithful-slice arms on ja/educational__244a3664's «・　夢として…» cell paragraph.

The synthetic probe (`_pb_cell_middot_lend_gen.py`) keeps 20 chars on line 1 under
compat 15 jc both like Oxi does, while the real document breaks after 19. Each arm
edits ONE thing in the real document.xml / settings.xml; readout = chars on the
paragraph's first line in Word's PDF.
"""
import os, sys, zipfile, glob, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/e244_dotline'; os.makedirs(OUT, exist_ok=True)
SRC = glob.glob('pipeline_data/docx_corpus/*/educational/244a36644e26b87a.docx')[0]
IND = '<w:ind w:leftChars="200" w:left="630" w:hangingChars="100" w:hanging="210"/>'
PARA = '・　夢としてはじめから就きたいものではなかったが'


def edit_para(doc, fn):
    for m in re.finditer(r'<w:p[ >].*?</w:p>', doc):
        seg = m.group(0)
        if PARA in ''.join(re.findall(r'<w:t[^>]*>([^<]*)', seg)):
            return doc[:m.start()] + fn(seg) + doc[m.end():]
    raise SystemExit('para not found')


ARMS = {
    'orig': (lambda d: d, lambda s: s),
    'no_ind': (lambda d: edit_para(d, lambda p: p.replace(IND, '')), lambda s: s),
    'twip_ind': (lambda d: edit_para(d, lambda p: p.replace(IND, '<w:ind w:left="630" w:hanging="210"/>')), lambda s: s),
    'jc_both': (lambda d: edit_para(d, lambda p: p.replace('<w:pPr>', '<w:pPr><w:jc w:val="both"/>', 1)), lambda s: s),
    'no_balance': (lambda d: d, lambda s: s.replace('<w:balanceSingleByteDoubleByteWidth/>', '')),
    'no_grid': (lambda d: re.sub(r'<w:docGrid [^>]*/>', '', d), lambda s: s),
    'no_dot': (lambda d: edit_para(d, lambda p: p.replace('・', '一', 1)), lambda s: s),
}

if __name__ == '__main__':
    import win32com.client, pymupdf
    z = zipfile.ZipFile(SRC); files = [(i.filename, z.read(i.filename)) for i in z.infolist()]
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for name, (fd, fs) in ARMS.items():
            at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as w:
                for fn, data in files:
                    if fn == 'word/document.xml':
                        data = fd(data.decode('utf8')).encode('utf8')
                    elif fn == 'word/settings.xml':
                        data = fs(data.decode('utf8')).encode('utf8')
                    w.writestr(fn, data)
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)[0]
            out = None
            for b in P.get_text('rawdict')['blocks']:
                for l in b.get('lines', []):
                    t = ''.join(c['c'] for s in l['spans'] for c in s['chars'])
                    if '夢としては' in t:
                        x0 = min(c['bbox'][0] for s in l['spans'] for c in s['chars'])
                        out = f'n {len(t.rstrip()):2} x0 {x0:.2f} {t.rstrip()}'
            print(f'{name:10} {out}', flush=True)
    finally:
        app.Quit()
