# -*- coding: utf-8 -*-
"""Controlled sweep: WHEN does Word wrap body text BESIDE a narrow floating

RESULTS (2026-07-06, sweeps 1+2):
- NO minimum remaining-width threshold: Word flows text into a strip as
  narrow as 30pt (right_rem30: 11 narrowed lines, ~29pt wide).
- OFFSET (tblpX) positioning == ALIGN (tblpXSpec) positioning (within 1pt).
- LEFT-side floats: text wraps on the RIGHT with an X-SHIFT (needs the
  band x-shift support, not just width reduction).
- Column OVERFLOW irrelevant: floats whose right edge exceeds content-right
  still get strip text (overflow_strip33 = the ed025c float0 twin: 3 strip
  lines) — the "off-column floats get no wrap" hypothesis is FALSE as a
  geometry rule.
- ★ed025c itself shows NO strip text in Word (16pp, 1 stray line) despite
  twin geometry → its cause is CONTENT-side (stacked forms/headings — no
  wrapping body text near the floats). What Oxi's banding moved (−0.0019)
  is UNRESOLVED; the S758b align-only gate stays until that is pinned.

table (vertAnchor=text), and at what remaining-width threshold?

Axes: float width (remaining usable width {30,60,90,120,180,240}pt) ×
positioning (tblpXSpec=right vs tblpX offset right-ish vs LEFT side) —
one config per page, body = numbered CJK paragraphs following the float.
Measure per config from the Word PDF: are the body lines beside the float
narrowed (side-wrap) or full-width below it (wrap-below / no-beside)?

Run: python tools/metrics/_sidewrap_sweep.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.environ.get("SIDEWRAP_DIR", os.path.join(os.environ.get("TEMP", "."), "sidewrap_sweep"))
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, "sidewrap_sweep.docx")
PDF = os.path.join(OUTDIR, "sidewrap_sweep.pdf")

SENT = pg.SENT
esc = pg.esc
CONTENT_W = 453.8  # A4 margins 1418tw both sides

def rpr():
    return (f'<w:rFonts w:ascii="{pg.MINCHO}" w:eastAsia="{pg.MINCHO}" w:hAnsi="{pg.MINCHO}"/>'
            '<w:sz w:val="21"/>')

def para(txt, brk=False):
    r = rpr()
    b = f'<w:r><w:rPr>{r}</w:rPr><w:br w:type="page"/></w:r>' if brk else ''
    return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>{b}'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

def float_tbl(w_tw, pos_attr, rows=6):
    r = rpr()
    cells = "".join(
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{w_tw}" w:type="dxa"/></w:tcPr>'
        f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">項{i+1}</w:t></w:r></w:p></w:tc></w:tr>'
        for i in range(rows))
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{w_tw}" w:type="dxa"/>'
            f'<w:tblpPr w:leftFromText="142" w:rightFromText="142" w:vertAnchor="text" {pos_attr} w:tblpY="1"/>'
            '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{w_tw}"/></w:tblGrid>' + cells + '</w:tbl>')

# remaining usable width = CONTENT_W − float_w − 7.1 (leftFromText)
configs = []
for rem_pt in [30, 60, 90, 120, 180, 240]:
    w_tw = int((CONTENT_W - rem_pt - 7.1) * 20)
    configs.append((f"right_rem{rem_pt}", w_tw, 'w:tblpXSpec="right"'))
# offset-positioned (right-ish): x chosen so the float's LEFT edge leaves rem on the left
for rem_pt in [90, 180]:
    w_tw = int((CONTENT_W - rem_pt - 7.1) * 20)
    x_tw = int((rem_pt + 7.1) * 20)
    configs.append((f"offs_rem{rem_pt}", w_tw, f'w:tblpX="{x_tw}"'))
# LEFT-side float (text should wrap on the RIGHT)
for rem_pt in [90, 180]:
    w_tw = int((CONTENT_W - rem_pt - 7.1) * 20)
    configs.append((f"left_rem{rem_pt}", w_tw, 'w:tblpXSpec="left"'))

body = []
for ci, (tag, w_tw, pos) in enumerate(configs):
    body.append(para(f"C{ci} 開始行。" + SENT[:20], brk=(ci > 0)))
    body.append(float_tbl(w_tw, pos))
    for j in range(6):
        body.append(para(f"第{ci}-{j}節　" + SENT))
pg.write_docx(DOCX, pg.doc(''.join(body) + pg.sectpr()))
print("wrote", DOCX, len(configs), "configs")

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
print("pages:", len(d))
for i, (tag, w_tw, pos) in enumerate(configs):
    if i >= len(d):
        break
    page = d[i]
    txt = page.get_text("dict")
    lines = []
    for b in txt["blocks"]:
        for l in b.get("lines", []):
            s = "".join(sp["text"] for sp in l.get("spans", []))
            if s.strip():
                lines.append((round(l["bbox"][1], 1), round(l["bbox"][0], 1),
                              round(l["bbox"][2], 1), s[:6]))
    lines.sort()
    # body lines (exclude the float's own 項N cells: short text at the float x)
    body_lines = [(y, x0, x1) for y, x0, x1, s in lines if not s.startswith("項")]
    narrowed = [(y, x0, x1) for y, x0, x1 in body_lines if (x1 - x0) < 400 and x1 < 500 and (x1 - x0) > 20]
    fulls = [(y, x0, x1) for y, x0, x1 in body_lines if x1 >= 500]
    right_side = [(y, x0, x1) for y, x0, x1 in body_lines if x0 > 100]
    print(f"{tag:14s} body={len(body_lines)} narrowed(right-cut)={len(narrowed)} full={len(fulls)} x-shifted={len(right_side)}"
          + (f"  eg_narrow={narrowed[1] if len(narrowed)>1 else narrowed[:1]}" if narrowed else "")
          + (f"  eg_shift={right_side[1] if len(right_side)>1 else right_side[:1]}" if right_side else ""))
