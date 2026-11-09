"""Faithful-slice probe: a paragraph whose last run ends with <w:br/> (zh dev creative__06787e0b:
«合肥-黄山<w:br/>» in a table cell; Word puts the next paragraph one whole line lower than Oxi).

Arms, from 06787e0b's package (its styles, its «travel» table style):
  body      : «甲<br/>» + «乙»
  body_nobr : «甲» + «乙»                     (control)
  cell      : the same two paragraphs inside a 1x1 «travel» table
  cell_nobr : control
  body_2br  : «甲<br/><br/>» + «乙»
Usage: python _pb_trailing_br_gen.py <src.docx> <out_dir>
Prints the baseline gap 甲 -> 乙 in Word's PDF and (if --oxi EXE) in Oxi's.
"""
import sys, os, re, zipfile

src, out = sys.argv[1], sys.argv[2]
oxi = sys.argv[sys.argv.index('--oxi') + 1] if '--oxi' in sys.argv else None
os.makedirs(out, exist_ok=True)
z = zipfile.ZipFile(src)
parts = {it.filename: z.read(it.filename) for it in z.infolist()}
doc = parts['word/document.xml'].decode('utf8')
head = doc[:doc.index('<w:body>') + len('<w:body>')]
sect = re.findall(r'<w:sectPr[ >].*?</w:sectPr>', doc, re.S)[-1]
sect = re.sub(r'<w:footerReference[^>]*/>', '', sect)


def p(text, brs=0):
    br = '<w:br/>' * brs
    return (f'<w:p><w:pPr><w:pStyle w:val="indent"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}{br}</w:t></w:r></w:p>')


def p_br(text, brs):
    br = '<w:br/>' * brs
    return (f'<w:p><w:pPr><w:pStyle w:val="indent"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t>{br}</w:r></w:p>')


def tbl(inner):
    return ('<w:tbl><w:tblPr><w:tblStyle w:val="travel"/></w:tblPr><w:tblGrid><w:gridCol w:w="9500"/></w:tblGrid>'
            f'<w:tr><w:trPr/><w:tc><w:tcPr><w:tcW w:w="9500" w:type="dxa"/></w:tcPr>{inner}</w:tc></w:tr></w:tbl>')


body2 = '<w:p><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>乙乙乙</w:t></w:r></w:p>'
ARMS = {
    'body': p_br('甲甲甲', 1) + body2,
    'body_nobr': p_br('甲甲甲', 0) + body2,
    'body_2br': p_br('甲甲甲', 2) + body2,
    'cell': tbl(p_br('甲甲甲', 1) + body2) + '<w:p/>',
    'cell_nobr': tbl(p_br('甲甲甲', 0) + body2) + '<w:p/>',
    # the document's own form: <w:br/> INSIDE <w:t> (generator output, not schema-valid)
    'body_brint': p('甲甲甲', 1) + body2,
    'cell_brint': tbl(p('甲甲甲', 1) + body2) + '<w:p/>',
}
import fitz, subprocess
import win32com.client
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False


def gap(pdf):
    pg = fitz.open(pdf)[0]
    y = {}
    for b in pg.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            t = ''.join(s['text'] for s in l['spans'])
            for k in ('甲', '乙'):
                if k in t and k not in y:
                    y[k] = l['spans'][0]['origin'][1]
    return round(y['乙'] - y['甲'], 2) if len(y) == 2 else None


try:
    for name, inner in ARMS.items():
        xml = head + inner + sect + '</w:body></w:document>'
        dst = os.path.join(out, name + '.docx')
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items():
                o.writestr(fn, xml if fn == 'word/document.xml' else raw)
        d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True)
        pdf = os.path.join(out, name + '.pdf')
        d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
        d.Close(False)
        line = f'{name:10s} Word {gap(pdf)}'
        if oxi:
            opdf = os.path.join(out, name + '.oxi.pdf')
            subprocess.run([oxi, 'docx-to-pdf', dst, opdf], capture_output=True)
            line += f'  Oxi {gap(opdf)}'
        print(line)
finally:
    w.Quit()
