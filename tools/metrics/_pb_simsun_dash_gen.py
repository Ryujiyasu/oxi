import zipfile, re, sys, os, fitz, subprocess
sys.stdout.reconfigure(encoding='utf-8')
S = sys.argv[1]; OXI = sys.argv[2]; os.makedirs(S + '/pb_dd', exist_ok=True)
z = zipfile.ZipFile('pipeline_data/docx_corpus/zh/reports/0075950ef0b1275c.docx')
parts = {i.filename: z.read(i.filename) for i in z.infolist()}
doc = parts['word/document.xml'].decode('utf8')
key = '1—7月份，我国电子信息制造业'
i = doc.index(key)
TXT = {'dash': '1—7月份', 'plain': '17月份', 'hyphen': '1-7月份', 'space_dash': '1 —7月份', 'cjkdash': '月—7月份', 'enfont_dash': None}
w = None
if '--word' in sys.argv:
    import win32com.client
    w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
try:
    for name, t in TXT.items():
        if t is None:
            continue
        xml = doc[:i] + t + doc[i + len('1—7月份'):]
        dst = f'{S}/pb_dd/{name}.docx'
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items(): o.writestr(fn, xml if fn == 'word/document.xml' else raw)
        res = {}
        if w:
            d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True); pdf = dst[:-5] + '.pdf'; d.ExportAsFixedFormat(os.path.abspath(pdf), 17); d.Close(False)
            res['word'] = pdf
        op = dst[:-5] + '.oxi.pdf'
        subprocess.run([OXI, 'docx-to-pdf', dst, op], capture_output=True, env=dict(os.environ, OXI_IGNORE_LRPB='1'))
        res['oxi'] = op
        out = []
        for k, pdf in res.items():
            for b in fitz.open(pdf)[0].get_text('rawdict')['blocks']:
                for l in b.get('lines', []):
                    cs = [(round(c['origin'][0], 2), c['c']) for s in l['spans'] for c in s['chars']]
                    s = ''.join(c for _, c in cs)
                    if '我国电子' in s:
                        j = s.index('月份')
                        out.append((k, [x for x in cs[:j + 1] if x[1] != ' ']))
        print(name, out)
finally:
    if w: w.Quit()
