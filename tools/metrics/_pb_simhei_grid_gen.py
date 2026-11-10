import zipfile, re, sys, os, fitz, subprocess
sys.stdout.reconfigure(encoding='utf-8')
S = sys.argv[1]; OXI = sys.argv[2]; os.makedirs(S + '/pb_snap', exist_ok=True)
z = zipfile.ZipFile('pipeline_data/docx_corpus/zh/creative/076f98c063ceb880.docx')
parts = {i.filename: z.read(i.filename) for i in z.infolist()}
doc = parts['word/document.xml'].decode('utf8')
body = doc.index('<w:body>') + 8
p1 = re.search(r'<w:p[ >].*?</w:p>', doc[body:], re.S)
a, b = body + p1.start(), body + p1.end(); P = p1.group(0)
ARMS = {}
for sz in (24, 28, 32, 36, 44):
    for line in ('360', '240', '480'):
        ARMS[f's{sz}_l{line}'] = P.replace('w:val="28"', f'w:val="{sz}"').replace('w:line="360"', f'w:line="{line}"')
ARMS['s28_l360_nobold'] = P.replace('<w:b/><w:bCs/>', '')
w = None
if '--word' in sys.argv:
    import win32com.client
    w = win32com.client.DispatchEx('Word.Application'); w.Visible = False; w.DisplayAlerts = 0


def base(pdf):
    pg = fitz.open(pdf)[0]
    ys = []
    for bl in pg.get_text('rawdict')['blocks']:
        for l in bl.get('lines', []):
            c = l['spans'][0]['chars'][0]
            t = ''.join(x['c'] for s in l['spans'] for x in s['chars'])
            if t.startswith('电子测量仪器的') or t.startswith('来源'):
                ys.append(round(c['origin'][1], 2))
    return ys


try:
    for name, p in ARMS.items():
        dst = f'{S}/pb_snap/{name}.docx'
        xml = doc[:a] + p + doc[b:]
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items(): o.writestr(fn, xml if fn == 'word/document.xml' else raw)
        res = {}
        if w:
            d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True); pdf = dst[:-5] + '.pdf'; d.ExportAsFixedFormat(os.path.abspath(pdf), 17); d.Close(False)
            res['word'] = base(pdf)
        op = dst[:-5] + '.oxi.pdf'
        subprocess.run([OXI, 'docx-to-pdf', dst, op], capture_output=True, env=dict(os.environ, OXI_IGNORE_LRPB='1'))
        res['oxi'] = base(op)
        print(name, res)
finally:
    if w: w.Quit()
