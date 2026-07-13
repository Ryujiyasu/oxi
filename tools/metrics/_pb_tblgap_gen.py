"""Latin paragraph -> TABLE-top gap derivation.

uklocalspending p36 (bundle state): Word's table top border sits +1.14pt
BELOW Oxi's model position (para el.y + line 12.649 + after 12 + bw 0.78
= 321.49 vs Word 322.63) while every row span inside the table is EXACT.
This probe derives Word's para->table gap rule directly (y readout, no
flip needed): an ANCHOR para (Arial 11, after=A) followed by either a
FOLLOWER para (control: gives Word's para->para gap = line + after) or a
2x2 TABLE (tcBorders sz, tcMar top=T, purple shading like uklocal).
Word PDF readout: anchor ink top, follower ink top / table top border y.

  delta(cfg) = (table border y) - (follower box top for same A)
             = the table-start offset vs a following paragraph.

Configs: A in {0, 240 direct, 240 style-level} x {follower, tbl sz6 T15,
tbl sz6 T105, tbl sz12 T15}.

Usage:
  python _pb_tblgap_gen.py gen
  python _pb_tblgap_gen.py measure
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_tblgap")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

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

# Normal with style-level after=240 only for the 'Asty' variants.
def styles(style_after):
    sp = '<w:spacing w:before="0" w:after="240"/>' if style_after else \
         '<w:spacing w:before="0" w:after="0"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:styles {W_NS}>'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
            f'<w:pPr><w:widowControl w:val="0"/>{sp}</w:pPr>'
            '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:style>'
            '</w:styles>')

R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'


def para(text, after_direct):
    if after_direct is None:
        ppr = f'<w:pPr><w:rPr>{R}</w:rPr></w:pPr>'  # style spacing applies
    else:
        ppr = (f'<w:pPr><w:spacing w:before="0" w:after="{after_direct}" '
               f'w:line="240" w:lineRule="auto"/><w:rPr>{R}</w:rPr></w:pPr>')
    return f'<w:p>{ppr}<w:r><w:rPr>{R}</w:rPr><w:t>{text}</w:t></w:r></w:p>'


def table(border_sz, tcmar_top):
    tc = []
    for txt in ('CellA', 'CellB'):
        tc.append(
            '<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>'
            f'<w:tcBorders><w:top w:val="single" w:sz="{border_sz}" w:space="0" w:color="000000"/>'
            f'<w:left w:val="single" w:sz="{border_sz}" w:space="0" w:color="000000"/>'
            f'<w:bottom w:val="single" w:sz="{border_sz}" w:space="0" w:color="000000"/>'
            f'<w:right w:val="single" w:sz="{border_sz}" w:space="0" w:color="000000"/></w:tcBorders>'
            '<w:shd w:val="clear" w:color="auto" w:fill="91278F"/>'
            f'<w:tcMar><w:top w:w="{tcmar_top}" w:type="dxa"/><w:left w:w="60" w:type="dxa"/>'
            f'<w:bottom w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tcMar></w:tcPr>'
            '<w:p><w:pPr><w:spacing w:before="120" w:after="120" w:line="0" w:lineRule="atLeast"/>'
            f'<w:rPr>{R}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R}<w:color w:val="FFFFFF"/></w:rPr><w:t>{txt}</w:t></w:r></w:p></w:tc>')
    row = f'<w:tr>{"".join(tc)}</w:tr>'
    row2 = row.replace('CellA', 'CellC').replace('CellB', 'CellD').replace('91278F', 'FFFFFF')
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblCellMar><w:top w:w="15" w:type="dxa"/><w:left w:w="15" w:type="dxa"/>'
            '<w:bottom w:w="15" w:type="dxa"/><w:right w:w="15" w:type="dxa"/></w:tblCellMar>'
            f'</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
            f'{row}{row2}</w:tbl>')


# cfg -> (anchor_after_direct(None=style), style_after, follower(True)/table(sz,T))
CFGS = {
    'pA0':   (0,    False, None),
    'pA240': (240,  False, None),
    'pAsty': (None, True,  None),
    'tA0':   (0,    False, (6, 15)),
    'tA240': (240,  False, (6, 15)),
    'tAsty': (None, True,  (6, 15)),
    'tT105': (240,  False, (6, 105)),
    'tSz12': (240,  False, (12, 15)),
    'tSz24': (240,  False, (24, 15)),
}


def build(cfg):
    after, style_after, tail = CFGS[cfg]
    body = ''.join(para(f'Filler {i} alpha beta.', 0) for i in range(3))
    body += para('ANCHORLINE omega.', after)
    if tail is None:
        body += para('FOLLOWERLINE psi.', 0)
    else:
        body += table(*tail)
        body += para('AFTERTBL chi.', 0)
    body += ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
             'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>'), styles(style_after)


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for cfg in CFGS:
        doc, sty = build(cfg)
        with zipfile.ZipFile(os.path.join(OUTDIR, f'ptg_{cfg}.docx'), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', doc)
            z.writestr('word/styles.xml', sty)
    print('generated', len(CFGS), 'docs in', OUTDIR)


def measure():
    import glob
    import win32com.client, fitz
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        for f in sorted(glob.glob(os.path.join(OUTDIR, 'ptg_*.docx'))):
            pdf = f[:-5] + '.pdf'
            if not os.path.exists(pdf):
                doc = word.Documents.Open(os.path.abspath(f), ReadOnly=True)
                doc.ExportAsFixedFormat(os.path.abspath(pdf), 17)
                doc.Close(False)
            d = fitz.open(pdf)
            pg = d[0]
            anchor = follower = None
            for blk in pg.get_text('dict')['blocks']:
                if blk.get('type') != 0:
                    continue
                for ln in blk['lines']:
                    t = ''.join(s['text'] for s in ln['spans'])
                    if 'ANCHORLINE' in t:
                        anchor = round(ln['bbox'][1], 2)
                    if 'FOLLOWERLINE' in t or 'CellA' in t:
                        follower = round(ln['bbox'][1], 2)
            # table top border: topmost thin BLACK rect / line below anchor
            border = None
            for dr in pg.get_drawings():
                fill = dr.get('fill')
                dark = fill is not None and max(fill) < 0.3
                for it in dr['items']:
                    y = None
                    if it[0] == 're' and it[1].height < 1.5 and it[1].width > 40 and dark:
                        y = it[1].y0
                    elif it[0] == 'l':
                        p1, p2 = it[1], it[2]
                        if abs(p1.y - p2.y) < 0.3 and abs(p1.x - p2.x) > 40:
                            y = p1.y
                    if y is not None and anchor and y > anchor:
                        border = y if border is None else min(border, y)
            base = os.path.basename(f)[:-5]
            print(f'{base}: anchor={anchor} next_ink={follower} tbl_border='
                  f'{round(border, 2) if border else None}')
    finally:
        word.Quit()


if __name__ == '__main__':
    if (sys.argv[1] if len(sys.argv) > 1 else 'gen') == 'gen':
        gen()
    else:
        measure()
