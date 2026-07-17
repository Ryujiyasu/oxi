# -*- coding: utf-8 -*-
"""Table-START orphan rule probe: does Word ever leave a table's FIRST row
alone at the page bottom when row 2 moves whole?

000bd832: Oxi placed row0 ('Essential requirements' header-ish, ~34pt) at the
p1 bottom and whole-moved row1 (trHeight-bound, 52pt) -> orphan; Word starts
the WHOLE table on p2. Sweep the table's start position across the boundary
with an exact spacer; read per-X which rows land on p1.

Shape: Letter, 1440 margins, Calibri 11. K fillers + exact spacer + a 4-row
table (row0 auto ~2 lines tall label row; rows 1-3 trHeight=1040tw atLeast).
Word truth via PDF (fitz): rows' first-cell texts per page.

Usage: gen [--sweep lo hi step] | measure
"""
import os, sys, zipfile, glob

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_tblstart")

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>'
          '</w:styles>')
FILLER = ('Lorem ipsum dolor sit amet consectetur adipiscing elit sed do '
          'eiusmod tempor.')


def cell(text, w=2340):
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>')


def build(name, spacer_tw):
    body = []
    for k in range(42):
        body.append(f'<w:p><w:r><w:t>F{k:02d} {FILLER}</w:t></w:r></w:p>')
    body.append(f'<w:p><w:pPr><w:spacing w:line="{spacer_tw}" w:lineRule="exact"/></w:pPr>'
                '<w:r><w:t>SPACER</w:t></w:r></w:p>')
    rows = []
    rows.append('<w:tr>' + cell('HEADROW essential requirements label') + cell('HEADCELL2') + cell('HEADCELL3') + cell('HEADCELL4') + '</w:tr>')
    for i in range(1, 4):
        rows.append(f'<w:tr><w:trPr><w:trHeight w:val="1040"/></w:trPr>'
                    + cell(f'DATA{i} content text') + cell(f'D{i}b') + cell(f'D{i}c') + cell(f'D{i}d') + '</w:tr>')
    body.append('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
                '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>'
                '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>'
                '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid>'
                + ''.join(rows) + '</w:tbl>')
    body.append('<w:p><w:r><w:t>AFTER paragraph.</w:t></w:r></w:p>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{"".join(body)}'
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(os.path.join(OUTDIR, name), 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOCRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)


def gen(sweep):
    os.makedirs(OUTDIR, exist_ok=True)
    for x in sweep:
        build(f'tbs_{x:05d}.docx', x)
    print('generated', len(sweep), 'docs in', os.path.abspath(OUTDIR))


def measure():
    import win32com.client
    import fitz
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for f in sorted(glob.glob(os.path.join(OUTDIR, 'tbs_*.docx'))):
            pdf = f[:-5] + '.pdf'
            if not os.path.exists(pdf):
                d = word.Documents.Open(os.path.abspath(f), ReadOnly=True)
                d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
                d.Close(False)
            doc = fitz.open(pdf)
            info = []
            for pi, pg in enumerate(doc):
                marks = []
                for blk in pg.get_text('dict')['blocks']:
                    for ln in blk.get('lines', []):
                        t = ''.join(s['text'] for s in ln['spans']).strip()
                        y = round(ln['bbox'][1], 1)
                        if t.startswith(('HEADROW', 'DATA', 'AFTER', 'SPACER')):
                            marks.append((y, t.split()[0]))
                marks.sort()
                info.append(f"p{pi+1}:" + ",".join(f"{m}@{y}" for y, m in marks))
            doc.close()
            print(os.path.basename(f)[:-5], ' | '.join(info))
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        if '--sweep' in sys.argv:
            i = sys.argv.index('--sweep')
            lo, hi, st = int(sys.argv[i+1]), int(sys.argv[i+2]), int(sys.argv[i+3])
            sweep = list(range(lo, hi + 1, st))
        else:
            sweep = [240, 400, 560, 720, 880, 1040, 1200]
        gen(sweep)
    else:
        measure()
