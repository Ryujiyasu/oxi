# -*- coding: utf-8 -*-
"""Controlled sweep: derive Word's tbRlV-cell AUTO row-height model (S753).

RESULT (2026-07-05): a tbRlV cell contributes ~ONE LINE to an auto row
regardless of n_chars (1..36), cell width (600..2400tw), fs (8/10.5), vAlign;
neighbor cells / trHeight always drive taller rows. The vertical text WRAPS
into the available row height (chars/col = floor((row_h-1)/fs)), columns
advance right-to-left, column pitch = the paragraph one-line height, block
horizontally centred in the cell, overflow past the border allowed.

Each config = one single-row table [vcell(tbRlV, W, n chars) | hcell(3000tw, text)]
separated by marker paragraphs. Measure row heights from the Word PDF border
lines (fitz get_drawings). Same doc skeleton as the probes (A4, docGrid lines
360, MS Mincho 10.5pt default, compat15).
"""
import os, sys
sys.path.insert(0, r"c:\Users\ryuji\oxi-main\tools\metrics")
import _probe_gen as pg
import _probe_gen3 as g3

SCRATCH = os.environ.get("TBRLV_SWEEP_DIR", os.path.join(os.environ.get("TEMP", "."), "tbrlv_sweep"))
os.makedirs(SCRATCH, exist_ok=True)
DOCX = os.path.join(SCRATCH, "tbrlv_sweep.docx")
PDF = os.path.join(SCRATCH, "tbrlv_sweep.pdf")

SENT = pg.SENT
esc = pg.esc

def rpr(sz="21"):
    return (f'<w:rFonts w:ascii="{pg.MINCHO}" w:eastAsia="{pg.MINCHO}" w:hAnsi="{pg.MINCHO}"/>'
            f'<w:sz w:val="{sz}"/>')

def vcell(txt, w, sz="21", valign=True, tdir=True):
    r = rpr(sz)
    va = '<w:vAlign w:val="center"/>' if valign else ''
    td = '<w:textDirection w:val="tbRlV"/>' if tdir else ''
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{td}{va}</w:tcPr>'
            f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p></w:tc>')

def hcell(txt, w=3000, sz="21"):
    r = rpr(sz)
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p></w:tc>')

def table(cells, trpr=""):
    total = 9000
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' + g3.pg2_borders() + '</w:tblPr>'
            '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="3000"/>' for _ in cells) + '</w:tblGrid>'
            f'<w:tr>{trpr}' + ''.join(cells) + '</w:tr></w:tbl>')

def marker(i):
    r = rpr()
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">M{i}</w:t></w:r></w:p>')

VT = SENT  # source text pool

configs = [
    # (tag, n_chars, W_tw, hcell_text, fs_sz, extra)
    ("n1_w1200",   1, 1200, "あ", "21", {}),
    ("n4_w1200",   4, 1200, "あ", "21", {}),
    ("n8_w1200",   8, 1200, "あ", "21", {}),
    ("n12_w1200", 12, 1200, "あ", "21", {}),
    ("n22_w1200", 22, 1200, "あ", "21", {}),
    ("n36_w1200", 36, 1200, "あ", "21", {}),
    ("n22_w600",  22,  600, "あ", "21", {}),
    ("n22_w2400", 22, 2400, "あ", "21", {}),
    ("n36_w2400", 36, 2400, "あ", "21", {}),
    ("n22_h2line",22, 1200, SENT[:60], "21", {}),   # hcell ~2-3 lines
    ("n22_h4line",22, 1200, "第1項：" + SENT, "21", {}),  # probe twin
    ("n8_fs8",     8, 1200, "あ", "16", {}),        # 8pt font
    ("n22_solo",  22, 1200, None, "21", {}),        # single-cell row
    ("n22_novalign",22,1200, "あ", "21", {"valign": False}),
    ("n22_trh50", 22, 1200, "あ", "21", {"trh": 1000}),  # trHeight atLeast 50pt
]

body_parts = []
mi = 0
for tag, n, w, htxt, sz, ex in configs:
    body_parts.append(marker(mi)); mi += 1
    cells = [vcell(VT[:n], w, sz, valign=ex.get("valign", True))]
    if htxt is not None:
        cells.append(hcell(htxt, sz=sz))
    trpr = ""
    if "trh" in ex:
        trpr = f'<w:trPr><w:trHeight w:val="{ex["trh"]}"/></w:trPr>'
    body_parts.append(table(cells, trpr))
body_parts.append(marker(mi))

body = "".join(body_parts) + pg.sectpr()
pg.write_docx(DOCX, pg.doc(body))
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
# collect horizontal border segments per page, in reading order
rows = []  # (page, y)
for pno in range(len(d)):
    page = d[pno]
    ys = set()
    for dr in page.get_drawings():
        for it in dr["items"]:
            if it[0] == "l":
                p1, p2 = it[1], it[2]
                if abs(p1.y - p2.y) < 0.2 and abs(p1.x - p2.x) > 30:
                    ys.add(round(p1.y, 2))
            elif it[0] == "re":
                r = it[1]
                if r.height < 1.5 and r.width > 30:
                    ys.add(round(r.y0, 2))
    for y in sorted(ys):
        rows.append((pno, y))

# pair consecutive border ys into tables (each single-row table = 2 lines).
# markers separate tables so lines come in clean pairs (dedup near-dups first)
clean = []
for p, y in rows:
    if clean and clean[-1][0] == p and abs(clean[-1][1] - y) < 1.0:
        continue
    clean.append((p, y))
print("\nborder lines:", clean)
print("\n-- row heights --")
i = 0
ci = 0
while i + 1 < len(clean) and ci < len(configs):
    p1, y1 = clean[i]; p2, y2 = clean[i + 1]
    if p1 == p2:
        tag = configs[ci][0]
        print(f"{tag:14s} row_h = {y2 - y1:7.2f}  (page {p1} y {y1:.2f}..{y2:.2f})")
        ci += 1
        i += 2
    else:
        # table split across pages?? shouldn't happen for these small rows
        print("PAGE SPLIT at", clean[i], clean[i + 1])
        i += 1
