"""Faithful-slice probe: where does a table's first-row text sit relative to the paragraph before
the table, as the table's top border width changes? (blind-J en/policies__00c94b4f: TableGrid sz 4,
Word's first-row baselines sit 0.25pt below Oxi's; body lines match.)

Arms: top border none / sz 4 / 8 / 12 / 24 (left/right/insideH kept at sz 4), plus a TableGrid
copy with all borders nil. Each arm = «Lead» paragraph + 1x1 table «Cell» + «After» paragraph.
Usage: python _pb_cell_top_border_gen.py <src.docx> <out_dir>
Prints per arm: Info(6) of Lead, Cell, After and the deltas; also exports probe.pdf so the
baselines can be read (span origin) next to Oxi's PDF.
"""
import sys, os, re, zipfile

src, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
ARMS = [('nil', 'nil', 0), ('s4', 'single', 4), ('s8', 'single', 8), ('s12', 'single', 12), ('s24', 'single', 24)]


def para(t):
    return f'<w:p><w:r><w:t>{t}</w:t></w:r></w:p>'


body = []
for name, val, sz in ARMS:
    top = f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="auto"/>' if val != 'nil' else '<w:top w:val="nil"/>'
    tbl = ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/>'
           f'<w:tblBorders>{top}</w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
           f'<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>{para("Cell " + name)}</w:tc></w:tr></w:tbl>')
    body += [para('Lead ' + name), tbl, para('After ' + name)]

z = zipfile.ZipFile(src)
doc = z.read('word/document.xml').decode('utf8')
head = doc[:doc.index('<w:body>') + len('<w:body>')]
sect = re.search(r'<w:sectPr[ >].*?</w:sectPr>', doc, re.S).group(0)
sect = re.sub(r'<w:(header|footer)Reference[^>]*/>', '', sect)
xml = head + ''.join(body) + sect + '</w:body></w:document>'
dst = os.path.join(out, 'probe.docx')
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        o.writestr(it, xml if it.filename == 'word/document.xml' else z.read(it.filename))

import win32com.client
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
try:
    d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True)
    ys = {}
    for i in range(1, d.Paragraphs.Count + 1):
        r = d.Paragraphs(i).Range
        t = r.Text.strip('\r\x07 ')
        ys[t] = d.Range(r.Start, r.Start).Information(6)
    for name, _, sz in ARMS:
        l, c, a = ys.get('Lead ' + name), ys.get('Cell ' + name), ys.get('After ' + name)
        print(name, sz, l, c, a, 'cell-lead', round(c - l, 2), 'after-cell', round(a - c, 2))
    d.ExportAsFixedFormat(os.path.abspath(os.path.join(out, 'probe.pdf')), 17)
    d.Close(False)
finally:
    w.Quit()
