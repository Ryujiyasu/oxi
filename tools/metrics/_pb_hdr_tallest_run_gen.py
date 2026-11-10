import zipfile, re, sys, os, fitz, win32com.client
sys.stdout.reconfigure(encoding='utf-8')
S = sys.argv[1]; os.makedirs(S + '/pb_hdrtall', exist_ok=True)
src = 'pipeline_data/docx_corpus/zh/legal/008491040f0dcb52.docx'
z = zipfile.ZipFile(src); parts = {i.filename: z.read(i.filename) for i in z.infolist()}
h = parts['word/header1.xml'].decode()
p1s = h.index('<w:p '); p1e = h.index('</w:p>', p1s) + 6
TNR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/></w:rPr>'
def para(runs):
    return '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>' + runs + '</w:p>'
def r(t, rpr=TNR):
    return f'<w:r>{rpr}<w:t xml:space="preserve">{t}</w:t></w:r>'
ARMS = {
    'orig': h[p1s:p1e],
    'latin_only': para(r('4/22/2021')),
    'cjk_only': para(r('审校日期')),
    'latin_big_cjk': para(r('AB') + r('审', '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="40"/></w:rPr>')),
    'calibri_cjk': para(r('审校日期', '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>')),
}
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False; w.DisplayAlerts = 0
try:
    for name, p1 in ARMS.items():
        hx = h[:p1s] + p1 + h[p1e:]
        dst = f'{S}/pb_hdrtall/{name}.docx'
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items(): o.writestr(fn, hx if fn == 'word/header1.xml' else raw)
        d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True)
        y = d.Paragraphs(2).Range.Information(6)
        pdf = dst[:-5] + '.pdf'; d.ExportAsFixedFormat(os.path.abspath(pdf), 17); d.Close(False)
        tb = None
        for b in fitz.open(pdf)[0].get_text('rawdict')['blocks']:
            for l in b.get('lines', []):
                t = ''.join(c['c'] for s in l['spans'] for c in s['chars'])
                if t.startswith('购房'): tb = round(l['spans'][0]['chars'][0]['origin'][1], 2)
        print(name, 'info6(title)', y, 'title baseline', tb)
finally:
    w.Quit()
