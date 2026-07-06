# -*- coding: utf-8 -*-
"""Controlled sweep: derive Word's table ROW-HEIGHT border/margin overhead
(the +0.5/row "border-box" cursor wall — probethdr −1×2, probeqtblstyle,
S200/S661/S666 render-only family, S375 tombstone).

Design: one docx, one 4-row 2-col table per config, marker paras between.
Word PDF border lines give per-row pitch. Algebra: within a config family,
pitch(4-line rows) − pitch(1-line rows) = 3 × cell_line_height (overhead
cancels); overhead = pitch(1-line) − line_height − padT − padB.

Sweep axes: border sz {0,4,8,12,24 eighth-pts} × cellMar t/b {0,60tw} ×
lines/cell {1,2,4} × fs {21,18 half-pts} × border SOURCE (table insideH vs
cell tcBorders vs none) × trHeight {none, atLeast 900, exact 900}.

Run: python tools/metrics/_rowbox_sweep.py  (writes docx+pdf to %TEMP%,
prints per-config row pitches).
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.environ.get("ROWBOX_DIR", os.path.join(os.environ.get("TEMP", "."), "rowbox_sweep"))
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, "rowbox_sweep.docx")
PDF = os.path.join(OUTDIR, "rowbox_sweep.pdf")

SENT = pg.SENT
esc = pg.esc

def rpr(sz="21"):
    return (f'<w:rFonts w:ascii="{pg.MINCHO}" w:eastAsia="{pg.MINCHO}" w:hAnsi="{pg.MINCHO}"/>'
            f'<w:sz w:val="{sz}"/>')

def cellp(txt, sz="21"):
    r = rpr(sz)
    return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

def borders(sz, inside=True):
    if sz == 0:
        return ''
    ih = f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>' if inside else ''
    return (f'<w:tblBorders><w:top w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>'
            f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>'
            f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>'
            f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>'
            f'{ih}<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/></w:tblBorders>')

def tc_borders(sz):
    b = f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>' \
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>'
    return f'<w:tcBorders>{b}</w:tcBorders>'

# text sized to wrap to n lines in a 3500tw (175pt) cell at 10.5pt:
# usable ~164pt -> ~15 fullwidth chars/line
def text_for(nlines, fs105=True):
    per = 15 if fs105 else 18
    n = per * nlines - 3
    base = (SENT * 3)
    return base[:n]

def table(tag, sz, cellmar_tb, nlines, fssz, source, trh):
    r = rpr(fssz)
    mar = (f'<w:tblCellMar><w:top w:w="{cellmar_tb}" w:type="dxa"/>'
           f'<w:left w:w="108" w:type="dxa"/><w:bottom w:w="{cellmar_tb}" w:type="dxa"/>'
           f'<w:right w:w="108" w:type="dxa"/></w:tblCellMar>') if cellmar_tb else ''
    tbl_borders = borders(sz, inside=True) if source == 'table' else (borders(0) if source == 'cell' else '')
    trpr = ''
    if trh == 'atleast':
        trpr = '<w:trPr><w:trHeight w:val="900"/></w:trPr>'
    elif trh == 'exact':
        trpr = '<w:trPr><w:trHeight w:val="900" w:hRule="exact"/></w:trPr>'
    rows = []
    for i in range(4):
        tcb = f'<w:tcPr><w:tcW w:w="3500" w:type="dxa"/>{tc_borders(sz) if source == "cell" else ""}</w:tcPr>'
        tcb2 = f'<w:tcPr><w:tcW w:w="5500" w:type="dxa"/>{tc_borders(sz) if source == "cell" else ""}</w:tcPr>'
        rows.append(f'<w:tr>{trpr}<w:tc>{tcb}{cellp(text_for(nlines), fssz)}</w:tc>'
                    f'<w:tc>{tcb2}{cellp("あ", fssz)}</w:tc></w:tr>')
    return (f'<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>{tbl_borders}{mar}</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3500"/><w:gridCol w:w="5500"/></w:tblGrid>'
            + ''.join(rows) + '</w:tbl>')

def marker(i):
    r = rpr()
    return (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:br w:type="page"/></w:r>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">M{i}</w:t></w:r></w:p>')

configs = [
    # (tag, sz, cellmar_tb, nlines, fssz, source, trh)
    ("sz4_4L",    4,  0, 4, "21", 'table', None),
    ("sz4_1L",    4,  0, 1, "21", 'table', None),
    ("sz4_2L",    4,  0, 2, "21", 'table', None),
    ("sz0_4L",    0,  0, 4, "21", 'none',  None),
    ("sz0_1L",    0,  0, 1, "21", 'none',  None),
    ("sz8_4L",    8,  0, 4, "21", 'table', None),
    ("sz8_1L",    8,  0, 1, "21", 'table', None),
    ("sz12_1L",  12,  0, 1, "21", 'table', None),
    ("sz24_1L",  24,  0, 1, "21", 'table', None),
    ("sz4mar_1L", 4, 60, 1, "21", 'table', None),
    ("sz4mar_4L", 4, 60, 4, "21", 'table', None),
    ("cell4_1L",  4,  0, 1, "21", 'cell',  None),
    ("cell4_4L",  4,  0, 4, "21", 'cell',  None),
    ("sz4_fs9_1L",4,  0, 1, "18", 'table', None),
    ("sz4_atl_1L",4,  0, 1, "21", 'table', 'atleast'),
    ("sz4_ex_1L", 4,  0, 1, "21", 'table', 'exact'),
]

body = [cellp("M0")]
for i, cfg in enumerate(configs):
    if i > 0:
        body.append(marker(i))
    body.append(table(*cfg))
pg.write_docx(DOCX, pg.doc(''.join(body) + pg.sectpr()))
print("wrote", DOCX)

import win32com.client
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
doc = word.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
doc.ExportAsFixedFormat(PDF, 17)
doc.Close(False)
word.Quit()
print("exported", PDF)

import fitz
d = fitz.open(PDF)
print("pages:", len(d), "configs:", len(configs))
for i, cfg in enumerate(configs):
    tag, sz, mar, nl, fssz, source, trh = cfg
    if i >= len(d):
        break
    page = d[i]
    ys = set()
    for dr in page.get_drawings():
        for it in dr["items"]:
            if it[0] == "l":
                p1, p2 = it[1], it[2]
                if abs(p1.y - p2.y) < 0.2 and abs(p1.x - p2.x) > 60:
                    ys.add(round(p1.y, 2))
            elif it[0] == "re":
                rr = it[1]
                if rr.height < 2.0 and rr.width > 60:
                    ys.add(round(rr.y0, 2))
    merged = []
    for y in sorted(ys):
        if merged and abs(merged[-1] - y) < 1.2:
            continue
        merged.append(y)
    pitches = [round(merged[j+1] - merged[j], 2) for j in range(len(merged) - 1)]
    print(f"{tag:12s} lines={len(merged)} pitches={pitches}")
