# -*- coding: utf-8 -*-
"""Does an anchored object in a vAlign=center cell count as cell content?

ja/administrative__1fc69a4f (blind-J p1): the calendar's notes cell (gridSpan 5,
vMerge restart over rows 274tw + 1455tw, vAlign center) holds ONE empty
paragraph whose runs anchor a 93pt wrapNone textbox (+7.3pt, layoutInCell) and
two pictures. Word keeps that paragraph at the cell top (its pictures at
paragraph +offset from 201.4 = the cell top); Oxi centred the one-line
paragraph in the 96pt span and the textbox moved 41pt down with it.

Faithful slice of that document (styles, theme, settings, sectPr, media): body
= the table's first three rows, a «■» run appended to the notes paragraph.
Arms vary the textbox height / presence. Readout: «■» y and the picture tops.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/administrative/1fc69a4f23b5918f.docx'
OUT = 'tests/fixtures/cell_anchor_valign'; os.makedirs(OUT, exist_ok=True)
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
      'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'v': 'urn:schemas-microsoft-com:vml'}
W = '{%s}' % NS['w']

# name -> (textbox height in pt or None = keep, keep_textbox, keep_pictures)
ARMS = {
    'orig': (None, True, True),
    'tb20': (20, True, True),
    'tb60': (60, True, True),
    'tb150': (150, True, True),
    'notb': (None, False, True),
    'noanch': (None, False, False),
}


def table(x, tb_h, keep_tb, keep_pics):
    from lxml import etree
    t = etree.fromstring(x.encode('utf8'))
    p = [p for p in t.iter(W + 'p') if 'はお' in ''.join(p.itertext())][0]
    p = p.xpath('ancestor::w:tc[1]/w:p[1]', namespaces=NS)[0]
    tbl = p.xpath('ancestor::w:tbl[1]', namespaces=NS)[0]
    for tr in tbl.findall(W + 'tr')[3:]:
        tbl.remove(tr)
    for r in p.findall(W + 'r'):
        is_tb = r.find('mc:AlternateContent', NS) is not None
        if (is_tb and not keep_tb) or (not is_tb and not keep_pics):
            p.remove(r)
        elif is_tb and tb_h is not None:
            emu = int(tb_h * 12700)
            for e in r.iter('{%s}extent' % NS['wp']):
                e.set('cy', str(emu))
            for e in r.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}ext'):
                if e.get('cy'):
                    e.set('cy', str(emu))
            for s in r.iter('{%s}rect' % NS['v']):
                s.set('style', re.sub(r'height:[\d.]+pt', f'height:{tb_h}pt', s.get('style', '')))
    p.append(etree.fromstring(
        f'<w:r xmlns:w="{NS["w"]}"><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" '
        f'w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr><w:t>■</w:t></w:r>'))
    return etree.tostring(tbl, encoding='unicode')


def build(name, arm):
    zin = zipfile.ZipFile(SRC)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, data in items:
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                tbl = table(x, *arm)
                b0 = x.index('<w:body>') + len('<w:body>')
                sect = re.search(r'<w:sectPr[ >].*?</w:sectPr>\s*</w:body>', x, re.S).group(0)
                x = x[:b0] + tbl + '<w:p/>' + sect + x[x.index('</w:body>') + len('</w:body>'):]
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    paths = {n: build(n, a) for n, a in ARMS.items()}
    if '--build' in sys.argv:
        sys.exit()
    import win32com.client, pymupdf
    for n, at in paths.items():
        app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(at[:-5] + '.pdf', 17)
            finally:
                d.Close(False)
        finally:
            app.Quit()
        page = pymupdf.open(at[:-5] + '.pdf')[0]
        sq = [round(s['bbox'][1], 1) for b in page.get_text('dict')['blocks'] for l in b.get('lines', [])
              for s in l['spans'] if '■' in s['text']]
        ims = sorted(round(i['bbox'][1], 1) for i in page.get_image_info())
        print(f'{n:8} ■ top {sq}  picture tops {ims}', flush=True)
