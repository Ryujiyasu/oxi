import zipfile, re, sys, os, fitz
sys.stdout.reconfigure(encoding='utf-8')
S = sys.argv[1]; os.makedirs(S + '/pb_ls', exist_ok=True)
z = zipfile.ZipFile('pipeline_data/docx_corpus/zh/reports/0075950ef0b1275c.docx')
parts = {i.filename: z.read(i.filename) for i in z.infolist()}
doc = parts['word/document.xml'].decode('utf8')
head = doc[:doc.index('<w:body>') + 8]; sect = re.findall(r'<w:sectPr[ >].*?</w:sectPr>', doc, re.S)[-1]
TEXTS = {
    # 4 gaps before the word + 3 punctuation
    'g4p3': '我国电子信息制造业产量6.54亿台，同比增长10.6%；微型计算机设备产量1.86亿台',
    # same gaps, punctuation -> ideographs
    'g4p0': '我国电子信息制造业产量6.54亿台和同比增长10.6%及微型计算机设备产量1.86亿台',
    # 2 gaps (6.54 -> ideographs), 3 punctuation
    'g2p3': '我国电子信息制造业产量六点五亿台，同比增长十点六；微型计算机设备产量1.86亿台',
    # 6 gaps
    'g6p3': '我国电子A信息B制造业产量6.54亿台，同比增长10.6%；微型计算机设备产量1.86亿台',
}
import win32com.client
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False; w.DisplayAlerts = 0


def first(pdf):
    for bl in fitz.open(pdf)[0].get_text('dict')['blocks']:
        for l in bl.get('lines', []):
            t = ''.join(s['text'] for s in l['spans']).replace(' ', '')
            if t.startswith('我国'):
                return t[-4:]


try:
    for name, text in TEXTS.items():
        res = []
        prev = None
        for right in range(0, 241, 6):
            rpr = '<w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/><w:sz w:val="24"/></w:rPr>'
            body = (f'<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:ind w:right="{right}"/>'
                    f'<w:jc w:val="left"/></w:pPr><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>')
            dst = f'{S}/pb_ls/{name}_{right}.docx'
            with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
                for fn, raw in parts.items():
                    o.writestr(fn, head + body + sect + '</w:body></w:document>' if fn == 'word/document.xml' else raw)
            d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True); pdf = dst[:-5] + '.pdf'
            d.ExportAsFixedFormat(os.path.abspath(pdf), 17); d.Close(False)
            end = first(pdf)
            if end != prev:
                res.append((right, end)); prev = end
        print(name, res, flush=True)
finally:
    w.Quit()
