import zipfile, re, sys, os, fitz, win32com.client
sys.stdout.reconfigure(encoding='utf-8')
S = sys.argv[1]; os.makedirs(S + '/pb_hlunit', exist_ok=True)
z = zipfile.ZipFile('pipeline_data/docx_corpus/zh/creative/076f98c063ceb880.docx')
parts = {i.filename: z.read(i.filename) for i in z.infolist()}
doc = parts['word/document.xml'].decode('utf8')
body = doc.index('<w:body>') + 8
ps = [m for m in re.finditer(r'<w:p[ >].*?</w:p>', doc[body:], re.S)]
p3 = ps[2]; a, b = body + p3.start(), body + p3.end(); P = p3.group(0)
hl = re.search(r'<w:hyperlink [^>]*>(.*?)</w:hyperlink>', P, re.S)
ARMS = {
    'orig': P,
    'nocache': P.replace(' w:firstLine="480"', ''),
    'unwrap': P.replace(hl.group(0), hl.group(1)),
    'unwrap_nocache': P.replace(hl.group(0), hl.group(1)).replace(' w:firstLine="480"', ''),
    'hl_sz16': P.replace(hl.group(0), hl.group(0).replace('<w:sz w:val="24"/>', '<w:sz w:val="32"/>')),
    'text_first': P.replace(hl.group(0), '<w:r><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/><w:sz w:val="24"/></w:rPr><w:t>前</w:t></w:r>' + hl.group(0)),
}
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False; w.DisplayAlerts = 0
try:
    for name, p in ARMS.items():
        xml = doc[:a] + p + doc[b:]
        dst = f'{S}/pb_hlunit/{name}.docx'
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items(): o.writestr(fn, xml if fn == 'word/document.xml' else raw)
        d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True); pdf = dst[:-5] + '.pdf'; d.ExportAsFixedFormat(os.path.abspath(pdf), 17); d.Close(False)
        x = None
        for bl in fitz.open(pdf)[0].get_text('dict')['blocks']:
            for l in bl.get('lines', []):
                t = ''.join(s['text'] for s in l['spans'])
                if t.startswith('电子测量仪器行业') or t.startswith('前电子'): x = round(l['bbox'][0] - 90, 2)
        print(name, 'indent', x)
finally:
    w.Quit()
