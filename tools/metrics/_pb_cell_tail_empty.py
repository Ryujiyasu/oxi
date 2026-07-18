# -*- coding: utf-8 -*-
"""Trailing EMPTY paragraphs in a table cell — Word's row-height rule.

policies__001cf65cd72d881c: the outer TableGrid data row carries 11
trailing empty paragraphs after the last text in the right cell; the row
bottom sits at Word slack ~41.5pt / Oxi ~53.3 (full-height empties would
be ~250) — BOTH engines collapse most of them, by different amounts
(+2.05 on the row = the S935-set default-ON knife edge). tokyoshugyo
wi=1529 recorded the same class ("Word collapses trailing auto empties
to ~0; clean-room repro failed to reproduce") — this matrix probes it
with the faithful ingredients: TableGrid style, docDefaults
after=160 line=259 sz=22, a 2-col row with text in both cells.

Matrix: N trailing empties in the right cell x {dd (docDefaults spacing
inherited), a0 (direct after=0 on the empties)}.
Readout: the table's bottom border y from the Word PDF drawings.

Usage: python _pb_cell_tail_empty.py gen | measure
"""
import os, sys, glob, shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_cell_tail_empty")

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>
<w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>
<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
<w:tblPr><w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders></w:tblPr></w:style>
</w:styles>"""


def p_text(t):
    return f'<w:p><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'


def p_empty(direct_a0):
    if direct_a0:
        return '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>'
    return '<w:p/>'


def build(n_empty, a0):
    cells_l = p_text('Left cell text one.') + p_text('Left cell text two.')
    cells_r = p_text('Right cell text.') + ''.join(p_empty(a0) for _ in range(n_empty))
    tbl = ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>'
           '<w:tr>'
           f'<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>{cells_l}</w:tc>'
           f'<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>{cells_r}</w:tc>'
           '</w:tr></w:tbl>')
    body = p_text('Anchor before.') + tbl + p_text('Anchor after.')
    body += pg.sectpr(pgsz='<w:pgSz w:w="12240" w:h="15840"/>',
                      mar='<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>',
                      grid='')
    return pg.doc(body)


CASES = [(n, a0) for n in [0, 1, 2, 3, 5, 11] for a0 in [False, True]]


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for n, a0 in CASES:
        nm = f'cte_n{n:02d}_{"a0" if a0 else "dd"}.docx'
        doc = build(n, a0)
        parts = {
            "[Content_Types].xml": pg.content_types(()),
            "_rels/.rels": pg.RELS,
            "word/document.xml": doc,
            "word/_rels/document.xml.rels": pg.docrels(()),
            "word/styles.xml": STYLES,
            "word/settings.xml": pg.settings_xml("15", False),
        }
        import zipfile
        with zipfile.ZipFile(os.path.join(OUTDIR, nm), 'w', zipfile.ZIP_DEFLATED) as z:
            for name, data in parts.items():
                z.writestr(name, data)
    print('generated', len(CASES))


def measure():
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for f in sorted(glob.glob(os.path.join(OUTDIR, '*.docx'))):
            pdf = f[:-5] + '.pdf'
            if os.path.exists(pdf):
                continue
            tmp = f[:-5] + '_t.docx'
            shutil.copy(f, tmp)
            doc = word.Documents.Open(os.path.abspath(tmp), ReadOnly=True)
            doc.ExportAsFixedFormat(os.path.abspath(pdf), 17)
            doc.Close(False)
            os.remove(tmp)
    finally:
        word.Quit()
    import fitz
    for f in sorted(glob.glob(os.path.join(OUTDIR, '*.pdf'))):
        pdf = fitz.open(f)
        pg1 = pdf[0]
        hs = []
        for dd in pg1.get_drawings():
            for it in dd['items']:
                if it[0] == 'l':
                    p1, p2 = it[1], it[2]
                    if abs(p1.y - p2.y) < 0.3 and abs(p1.x - p2.x) > 40:
                        hs.append(round(p1.y, 2))
                elif it[0] == 're':
                    r = it[1]
                    if r.height < 1.2 and r.width > 40:
                        hs.append(round(r.y0, 2))
        hs = sorted(set(hs))
        # last anchor baseline
        anchor = None
        for b in pg1.get_text('dict')['blocks']:
            for l in b.get('lines', []):
                t = ''.join(s['text'] for s in l['spans']).strip()
                if t == 'Anchor after.':
                    anchor = round(l['spans'][0]['origin'][1], 2)
        print(f'{os.path.basename(f)}: borders {hs} anchor_after {anchor}')


if __name__ == '__main__':
    {'gen': gen, 'measure': measure}[sys.argv[1]]()
