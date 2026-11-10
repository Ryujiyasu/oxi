import zipfile, re, sys, os, fitz, subprocess
sys.stdout.reconfigure(encoding='utf-8')
S = sys.argv[1]; OXI = sys.argv[2]; os.makedirs(S + '/pb_ce', exist_ok=True)
z = zipfile.ZipFile('pipeline_data/docx_corpus/zh/technical/014c47f07d612993.docx')
parts = {i.filename: z.read(i.filename) for i in z.infolist()}
doc = parts['word/document.xml'].decode('utf8'); st = parts['word/styles.xml'].decode('utf8')
EMPTY = '<w:ind w:firstLine="480"/><w:jc w:val="left"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:kern w:val="0"/><w:sz w:val="24"/></w:rPr></w:pPr></w:p>'
ARMS = {}
for line in ('240', '360', '480'):
    for sz in ('24', '36', '44'):
        ARMS[f'l{line}_s{sz}'] = (doc.replace(EMPTY, EMPTY.replace('w:val="24"', f'w:val="{sz}"')),
                                   st.replace('<w:spacing w:line="360" w:lineRule="auto"/><w:ind w:firstLineChars="200"', f'<w:spacing w:line="{line}" w:lineRule="auto"/><w:ind w:firstLineChars="200"'))
w = None
if '--word' in sys.argv:
    import win32com.client
    w = win32com.client.DispatchEx('Word.Application'); w.Visible = False; w.DisplayAlerts = 0


def rows(pdf):
    ys = []
    for bl in fitz.open(pdf)[0].get_text('rawdict')['blocks']:
        for l in bl.get('lines', []):
            t = ''.join(x['c'] for s in l['spans'] for x in s['chars'])
            if t.startswith('酵母粉') or t.startswith('蛋白胨') or t.startswith('葡萄糖'):
                ys.append(round(l['spans'][0]['chars'][0]['origin'][1], 2))
    return [round(b - a, 2) for a, b in zip(ys, ys[1:])]


try:
    for name, (d, s) in ARMS.items():
        dst = f'{S}/pb_ce/{name}.docx'
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items():
                o.writestr(fn, d if fn == 'word/document.xml' else s if fn == 'word/styles.xml' else raw)
        res = {}
        if w:
            dd = w.Documents.Open(os.path.abspath(dst), ReadOnly=True); pdf = dst[:-5] + '.pdf'; dd.ExportAsFixedFormat(os.path.abspath(pdf), 17); dd.Close(False)
            res['word'] = rows(pdf)
        op = dst[:-5] + '.oxi.pdf'
        subprocess.run([OXI, 'docx-to-pdf', dst, op], capture_output=True, env=dict(os.environ, OXI_IGNORE_LRPB='1'))
        res['oxi'] = rows(op)
        print(name, res)
finally:
    if w: w.Quit()
