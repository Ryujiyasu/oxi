import zipfile, re, sys, os, fitz, subprocess
sys.stdout.reconfigure(encoding='utf-8')
S = sys.argv[1]; OXI = sys.argv[2]; os.makedirs(S + '/pb_cm', exist_ok=True)
z = zipfile.ZipFile('pipeline_data/docx_corpus/zh/forms/004fbbfd4b2c27c3.docx')
parts = {i.filename: z.read(i.filename) for i in z.infolist()}
doc = parts['word/document.xml'].decode('utf8'); num = parts['word/numbering.xml'].decode('utf8')
head = doc[:doc.index('<w:body>') + 8]; sect = re.findall(r'<w:sectPr[ >].*?</w:sectPr>', doc, re.S)[-1]


def cell(inner, w=2000):
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{inner}</w:tc>'


def para(text, num=True, before=240):
    n = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>' if num else ''
    r = f'<w:r><w:t>{text}</w:t></w:r>' if text else ''
    return f'<w:p><w:pPr>{n}<w:spacing w:before="{before}" w:line="276" w:lineRule="auto"/><w:ind w:firstLine="0"/></w:pPr>{r}</w:p>'


def table(c2):
    rows = ''.join(f'<w:tr>{cell(para("R" + str(i), False))}{cell(c2)}</w:tr>' for i in range(4))
    return ('<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' + rows + '</w:tbl><w:p/>')


def numsz(sz):
    return re.sub(r'(<w:abstractNum [^>]*w:abstractNumId="0".*?<w:lvl w:ilvl="0".*?<w:sz w:val=")\d+(")', lambda m: m.group(1) + str(sz) + m.group(2), num, count=1, flags=re.S)


ARMS = {}
for sz in (22, 36, 52):
    ARMS[f'empty_s{sz}'] = (table(para('', True)), numsz(sz))
    ARMS[f'text_s{sz}'] = (table(para('Q', True)), numsz(sz))
ARMS['nolist'] = (table(para('', False)), num)
w = None
if '--word' in sys.argv:
    import win32com.client
    w = win32com.client.DispatchEx('Word.Application'); w.Visible = False; w.DisplayAlerts = 0


def rows(pdf):
    ys = []
    for bl in fitz.open(pdf)[0].get_text('rawdict')['blocks']:
        for l in bl.get('lines', []):
            t = ''.join(x['c'] for s in l['spans'] for x in s['chars'])
            if re.match(r'R\d', t):
                ys.append(round(l['spans'][0]['chars'][0]['origin'][1], 2))
    return [round(b - a, 2) for a, b in zip(ys, ys[1:])][:2]


try:
    for name, (body, nx) in ARMS.items():
        dst = f'{S}/pb_cm/{name}.docx'
        xml = head + body + sect + '</w:body></w:document>'
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items():
                o.writestr(fn, xml if fn == 'word/document.xml' else nx if fn == 'word/numbering.xml' else raw)
        res = {}
        if w:
            d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True); pdf = dst[:-5] + '.pdf'; d.ExportAsFixedFormat(os.path.abspath(pdf), 17); d.Close(False)
            res['word'] = rows(pdf)
        op = dst[:-5] + '.oxi.pdf'
        subprocess.run([OXI, 'docx-to-pdf', dst, op], capture_output=True, env=dict(os.environ, OXI_IGNORE_LRPB='1'))
        res['oxi'] = rows(op)
        print(name, res)
finally:
    if w: w.Quit()
