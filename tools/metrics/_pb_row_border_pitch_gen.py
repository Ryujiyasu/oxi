"""Faithful-slice probe: does a cell's border add to an atLeast row's height? (zh dev
administrative__003ed1b3: rows trHeight 402 (20.1pt) + cell margins 57tw, cells with only a
bottom single sz4 border; Word's row pitch 25.8 = 20.1 + 5.7, Oxi 26.3.)

Arms (3-row, 2-column table, the document's cell margins, lines grid, SimSun 11pt):
  bot4 (the document), none, bot12, topbot4, top4, all4 (tblBorders single 4)
Row pitch = baseline gap between rows 1->2 and 2->3 in Word's PDF; --oxi EXE gives Oxi's too.
Usage: python _pb_row_border_pitch_gen.py <src.docx> <out_dir> [--oxi EXE]
"""
import sys, os, re, zipfile, subprocess

sys.stdout.reconfigure(encoding='utf-8')
src, out = sys.argv[1], sys.argv[2]
oxi = sys.argv[sys.argv.index('--oxi') + 1] if '--oxi' in sys.argv else None
os.makedirs(out, exist_ok=True)
z = zipfile.ZipFile(src)
parts = {it.filename: z.read(it.filename) for it in z.infolist()}
doc = parts['word/document.xml'].decode('utf8')
head = doc[:doc.index('<w:body>') + len('<w:body>')]
sect = re.findall(r'<w:sectPr[ >].*?</w:sectPr>', doc, re.S)[-1]
sect = re.sub(r'<w:(header|footer)Reference[^>]*/>', '', sect)

B = lambda side, sz: f'<w:{side} w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>' if sz else f'<w:{side} w:val="nil"/>'
ARMS = {'bot4': (0, 4), 'none': (0, 0), 'bot12': (0, 12), 'topbot4': (4, 4), 'top4': (4, 0)}
# --grid: trHeight x tcMar x bottom border, to find when a bottom border takes height
GRID = '--grid' in sys.argv


def table(top, bot, tblb='', trh=402, mar=57):
    rows = ''
    for k in range(3):
        cells = ''
        for w, t in ((1000, f'{k + 1}'), (5000, f'第{k + 1}行测试文字')):
            cells += (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/><w:tcBorders>{B("top", top)}{B("left", 4)}{B("bottom", bot)}{B("right", 0)}</w:tcBorders>'
                      f'<w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:widowControl/><w:rPr><w:sz w:val="22"/></w:rPr></w:pPr>'
                      f'<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>{t}</w:t></w:r></w:p></w:tc>')
        trp = f'<w:trHeight w:val="{trh}"/>' if trh else ''
        rows += f'<w:tr><w:trPr>{trp}<w:jc w:val="center"/></w:trPr>{cells}</w:tr>'
    return ('<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:jc w:val="center"/>' + tblb +
            f'<w:tblCellMar><w:top w:w="{mar}" w:type="dxa"/><w:left w:w="57" w:type="dxa"/><w:bottom w:w="{mar}" w:type="dxa"/>'
            '<w:right w:w="57" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="5000"/></w:tblGrid>'
            + rows + '</w:tbl>')


import fitz, win32com.client


def pitches(pdf):
    pg = fitz.open(pdf)[0]
    ys = {}
    for b in pg.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            t = ''.join(s['text'] for s in l['spans']).replace(' ', '')
            for k in (1, 2, 3):
                if f'第{k}行' in t:
                    ys[k] = l['spans'][0]['origin'][1]
    return [round(ys[k + 1] - ys[k], 2) for k in (1, 2)] if len(ys) == 3 else None


w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
try:
    arms = ({f'trh{t}_mar{m}_bot{b}': (0, b, t, m) for t in (0, 312, 402, 600) for m in (0, 57) for b in (0, 4)}
            if GRID else {f'mar{m}_bot{b}': (0, b, 0, m) for m in (5, 10, 20) for b in (0, 12, 24)}
            if '--partial' in sys.argv else {k: (v[0], v[1], 402, 57) for k, v in ARMS.items()})
    for name, (top, bot, trh, mar) in arms.items():
        xml = head + '<w:p/>' + table(top, bot, '', trh, mar) + '<w:p/>' + sect + '</w:body></w:document>'
        dst = os.path.join(out, name + '.docx')
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
            for fn, raw in parts.items():
                o.writestr(fn, xml if fn == 'word/document.xml' else raw)
        d = w.Documents.Open(os.path.abspath(dst), ReadOnly=True)
        pdf = os.path.join(out, name + '.pdf')
        d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
        d.Close(False)
        line = f'{name:8s} Word {pitches(pdf)}'
        if oxi:
            opdf = os.path.join(out, name + '.oxi.pdf')
            subprocess.run([oxi, 'docx-to-pdf', dst, opdf], capture_output=True)
            line += f'  Oxi {pitches(opdf)}'
        print(line)
finally:
    w.Quit()
