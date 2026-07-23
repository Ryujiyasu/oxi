# -*- coding: utf-8 -*-
"""R1: does the BODY's page-bottom limit track the footer's GLYPH-INK top
(H-ink) or its LINE-BOX top (H-box, current Oxi)?

forms__001c51d6: Oxi rejects an empty atLeast bordered terminal row's bottom
border ~2.13pt early (row height correct to 0.07pt) because it uses the footer
line-box top as the body hard bottom; Word lets the border descend to ~0.68pt
above the footer's first glyph ink (§3.2 of REPORT).

METHOD: body = K filler lines + a terminal EMPTY atLeast single-cell BORDERED
row (trHeight H). footer = 1-row 2-cell text-bearing table (text size S) +
trailing empty Footer paragraph. A MARKER paragraph sits INSIDE the terminal
row's cell as the LAST line so its page is observable via COM (empty rows have
no text to query). Actually: the row is the last body block; put a 1-char
marker paragraph AFTER the table so we can read which page the table ended on,
AND read the row's bottom border + footer ink from the Word PDF.

Sweep a preceding exact spacer (line=X exact, 2tw steps) to move the terminal
row's start across the page-bottom boundary; the flip (row bottom border on
p1 vs p2) pins Word's body-bottom limit to 0.1pt. Vary footer text size S:
  H-ink  -> flip tracks footer glyph-ink top -> moves with S
  H-box  -> flip tracks footer line-box top

Geometry: A4 (595x842pt), top 36pt, bottom 36pt, footer dist 35.4pt (708tw).
Arial-11 line 12.6489pt.

Usage: python _pb_r1_footer_collision_gen.py gen | measure
"""
import os, sys, zipfile, json
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                   "pipeline_data", "_pb_r1")
WNS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
         '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {WNS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/>'
          '<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:style>'
          '</w:styles>')
R11 = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
SP0 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'


def footer_xml(sz_half):
    rf = f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="{sz_half}"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:ftr {WNS}>'
            '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>'
            '<w:tr>'
            f'<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr>{SP0}<w:rPr>{rf}</w:rPr></w:pPr><w:r><w:rPr>{rf}</w:rPr><w:t>FTRLEFT</w:t></w:r></w:p></w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr>{SP0}<w:rPr>{rf}</w:rPr></w:pPr><w:r><w:rPr>{rf}</w:rPr><w:t>1</w:t></w:r></w:p></w:tc>'
            '</w:tr></w:tbl>'
            f'<w:p><w:pPr><w:pStyle w:val="Footer"/>{SP0}<w:rPr>{R11}</w:rPr></w:pPr></w:p>'
            '</w:ftr>')


def build(k, spacer_tw, trh_tw, terminal='empty'):
    fill = ''.join(
        f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{R11}</w:rPr><w:t>F{i:02d} filler alpha beta gamma.</w:t></w:r></w:p>'
        for i in range(k))
    spacer = (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{spacer_tw}" '
              f'w:lineRule="exact"/><w:rPr>{R11}</w:rPr></w:pPr></w:p>') if spacer_tw else ''
    if terminal == 'empty':
        cell_p = f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr></w:p>'
    else:
        cell_p = (f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr>'
                  f'<w:r><w:rPr>{R11}</w:rPr><w:t>ROWTEXT content line.</w:t></w:r></w:p>')
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
           '<w:tblBorders>'
           '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '</w:tblBorders></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
           f'<w:tr><w:trPr><w:trHeight w:val="{trh_tw}"/></w:trPr>'
           f'<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{cell_p}</w:tc></w:tr></w:tbl>')
    marker = (f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr>'
              f'<w:r><w:rPr>{R11}</w:rPr><w:t>TAILMARK</w:t></w:r></w:p>')
    body = (fill + spacer + tbl + marker +
            '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
            'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {WNS}><w:body>{body}</w:body></w:document>')


# terminal empty atLeast row trH=2079 (103.95pt, forms row35). K=48 fillers put
# the row near the p1 bottom; sweep spacer to cross the boundary. footer sz in
# half-points: 16=8pt, 20=10pt, 22=11pt, 24=12pt.
CASES = []
for sz in (16, 22, 24):              # footer 8 / 11 / 12 pt
    for sp in range(0, 252, 12):     # spacer 0..240 tw (0.6pt steps) brackets the row flip
        CASES.append((f"ink_s{sz}_sp{sp:03d}", 50, sp, 2079, 'empty', sz))


def gen():
    os.makedirs(OUT, exist_ok=True)
    for nm, k, sp, trh, term, sz in CASES:
        with zipfile.ZipFile(os.path.join(OUT, nm + '.docx'), 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DRELS)
            z.writestr('word/document.xml', build(k, sp, trh, term))
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/footer1.xml', footer_xml(sz))
    print(f'generated {len(CASES)} -> {os.path.abspath(OUT)}')


def measure():
    import win32com.client, fitz
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False; word.DisplayAlerts = 0
    res = {}
    try:
        for nm, k, sp, trh, term, sz in CASES:
            p = os.path.abspath(os.path.join(OUT, nm + '.docx'))
            pdf = os.path.abspath(os.path.join(OUT, nm + '.pdf'))
            d = word.Documents.Open(p, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)  # wdExportFormatPDF
            finally:
                d.Close(False)
            doc = fitz.open(pdf)
            # footer ink top on page 1 (FTRLEFT)
            ftr_ink = None
            for b in doc[0].get_text('dict')['blocks']:
                for ln in b.get('lines', []):
                    if 'FTRLEFT' in ''.join(s['text'] for s in ln['spans']):
                        ftr_ink = round(ln['bbox'][1], 2)
            # find the TABLE bottom border and which page it is on. The table has
            # a full box (top+bottom borders ~104pt apart). Detect it per page.
            def hlines(pg):
                ys = []
                for dr in pg.get_drawings():
                    for it in dr['items']:
                        if it[0] == 'l' and abs(it[1].y - it[2].y) < 0.4:
                            ys.append(round(it[1].y, 2))
                        elif it[0] == 're':
                            ys.append(round(it[1].y0, 2)); ys.append(round(it[1].y1, 2))
                return sorted(ys)
            row_page = None; row_bottom = None
            for pno in range(doc.page_count):
                ys = [y for y in hlines(doc[pno]) if y < 800]
                if len(ys) >= 2 and (max(ys) - min(ys)) > 90:  # the ~104pt table box
                    row_page = pno + 1; row_bottom = max(ys)
            # marker page (table stays on p1 even when marker overflows to p2)
            mk_pg = None
            for pno in range(doc.page_count):
                for b in doc[pno].get_text('dict')['blocks']:
                    for ln in b.get('lines', []):
                        if 'TAILMARK' in ''.join(s['text'] for s in ln['spans']):
                            mk_pg = pno + 1
            res[nm] = {'sz': sz, 'sp': sp, 'npages': doc.page_count,
                       'row_page': row_page, 'row_bottom': row_bottom,
                       'footer_ink_top': ftr_ink, 'mk_pg': mk_pg}
            doc.close()
            print(f"  {nm}: row_page {row_page} row_bottom {row_bottom} ftr_ink {ftr_ink} mk p{mk_pg}")
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUT, '_result.json'), 'w'), indent=1)
    print("wrote _result.json")


if __name__ == '__main__':
    (gen if (len(sys.argv) > 1 and sys.argv[1] == 'gen') else measure)()
