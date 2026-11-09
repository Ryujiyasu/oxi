"""Faithful-slice probe: what does Word add to an hRule=exact row -- the TABLE bottom margin or the
CELL one? (zh dev technical__006f0290 footer bar: exact 115tw, tblCellMar bottom 144, cells tcMar 0.)
Usage: python _pb_exact_row_cellmar_gen.py <out_dir>   (run from the repo root)
"""
import zipfile, re, sys, os, fitz, win32com.client
S = sys.argv[1]; os.makedirs(S + '/pb_exrow', exist_ok=True)
z = zipfile.ZipFile('pipeline_data/docx_corpus/zh/technical/006f02904419bd79.docx')
parts = {i.filename: z.read(i.filename) for i in z.infolist()}
doc = parts['word/document.xml'].decode('utf8')
head = doc[:doc.index('<w:body>') + 8]; sect = re.findall(r'<w:sectPr[ >].*?</w:sectPr>', doc, re.S)[-1]
def tbl(tblb, cellb):
    tcm = f'<w:tcMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="{cellb}" w:type="dxa"/></w:tcMar>' if cellb is not None else ''
    row = (f'<w:tr><w:trPr><w:trHeight w:val="1200" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>'
           f'<w:shd w:val="clear" w:color="auto" w:fill="0070C0"/>{tcm}</w:tcPr><w:p/></w:tc></w:tr>')
    return (f'<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="{tblb}" w:type="dxa"/>'
            f'</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>{row}</w:tbl>')
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
try:
    for name, (tb, cb) in {'tbl200_cellNone': (200, None), 'tbl200_cell0': (200, 0), 'tbl200_cell400': (200, 400), 'tbl0_cell400': (0, 400)}.items():
        xml = head + '<w:p><w:r><w:t>Top</w:t></w:r></w:p>' + tbl(tb, cb) + '<w:p/>' + sect + '</w:body></w:document>'
        dst = f'{S}/pb_exrow/{name}.docx'
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items(): o.writestr(fn, xml if fn == 'word/document.xml' else raw)
        d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True); pdf = dst[:-5] + '.pdf'; d.ExportAsFixedFormat(os.path.abspath(pdf), 17); d.Close(False)
        r = [dr['rect'] for dr in fitz.open(pdf)[0].get_drawings() if dr.get('fill') and dr['rect'].width > 100]
        print(name, 'row height', round(r[0].height, 2) if r else None)
finally:
    w.Quit()
