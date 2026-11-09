import zipfile, re, sys, os, fitz, win32com.client
S = sys.argv[1]; os.makedirs(S + '/pb_sectnum', exist_ok=True)
z = zipfile.ZipFile('pipeline_data/docx_corpus/zh/technical/006f02904419bd79.docx')
parts = {i.filename: z.read(i.filename) for i in z.infolist()}
doc = parts['word/document.xml'].decode('utf8')
head = doc[:doc.index('<w:body>') + 8]; sect = re.findall(r'<w:sectPr[ >].*?</w:sectPr>', doc, re.S)[-1]
mid_sect = '<w:sectPr><w:pgSz w:w="11900" w:h="16840"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="851" w:footer="421" w:gutter="0"/></w:sectPr>'
P = lambda t: f'<w:p><w:pPr><w:pStyle w:val="38"/></w:pPr><w:r><w:t>{t}</w:t></w:r></w:p>'
ARMS = {
    'empty_sect': P('AAA') + f'<w:p><w:pPr><w:pStyle w:val="38"/>{mid_sect}</w:pPr></w:p>' + P('BBB'),
    'empty_plain': P('AAA') + '<w:p><w:pPr><w:pStyle w:val="38"/></w:pPr></w:p>' + P('BBB'),
    'text_sect': P('AAA') + f'<w:p><w:pPr><w:pStyle w:val="38"/>{mid_sect}</w:pPr><w:r><w:t>XXX</w:t></w:r></w:p>' + P('BBB'),
}
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
try:
    for name, body in ARMS.items():
        xml = head + body + sect + '</w:body></w:document>'
        dst = f'{S}/pb_sectnum/{name}.docx'
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items(): o.writestr(fn, xml if fn == 'word/document.xml' else raw)
        d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True); pdf = dst[:-5] + '.pdf'; d.ExportAsFixedFormat(os.path.abspath(pdf), 17); d.Close(False)
        txt = ' | '.join(pg.get_text().replace('\n', ' ').strip()[:80] for pg in fitz.open(pdf))
        print(name, 'WORD:', txt)
finally:
    w.Quit()
