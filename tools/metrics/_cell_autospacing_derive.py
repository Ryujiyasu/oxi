# -*- coding: utf-8 -*-
"""Derive Word's before/afterAutospacing amount INSIDE A TABLE CELL (Ra, 2026-07-01).

S675 derived the BODY value (constant 13.75pt, direct-only). The cell content path
(mod.rs) does NOT apply before/after_autospacing at all → Oxi reserves 0 in cells.
This measures, via Word COM Information(6) per-paragraph Y (R30 collapsed-start fix):
  (1) BODY sanity  — afterAutospacing on a body para  (expect 13.75)
  (2) CELL value   — after/beforeAutospacing on an INTERIOR cell para
  (3) CELL EDGE    — first-para before-auto / last-para after-auto suppression
  (4) collapse     — adjacent autospacing cell paras
  (5) size sweep   — is the cell value constant like the body?
"""
import os, sys, io
sys.path.insert(0, 'tools/metrics')
import win32com.client
from mixedh_lineplace import build_generic
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')

FONT = 'ＭＳ 明朝'

def rpr(sz=22):
    return ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/>'
            % (FONT, FONT, FONT, sz))

def para(text, sz=22, spacing_attrs=''):
    sp = ('<w:spacing %s/>' % spacing_attrs) if spacing_attrs else ''
    return ('<w:p><w:pPr>%s<w:rPr>%s</w:rPr></w:pPr>'
            '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (sp, rpr(sz), rpr(sz), text))

def cell(inner_paras):
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '</w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>'
            '%s</w:tc></w:tr></w:tbl>' % inner_paras)

def ys(word, docx):
    """Per-paragraph (Information(6) Y, text) with R30 collapsed-start fix."""
    doc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
    out = []
    try:
        for p in doc.Paragraphs:
            rng = p.Range
            sr = doc.Range(rng.Start, rng.Start)
            out.append((sr.Information(6), p.Range.Text.strip()))
    finally:
        doc.Close(False)
    return out

def gap_after(word, name, body):
    """gap MID->BOTTOM for the autospaced doc minus its normal control = afterAuto."""
    y = ys(word, build_generic(name, body))
    return y

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        # ---- (1) BODY sanity: afterAutospacing on a body para (expect ~13.75) ----
        print("=== (1) BODY sanity (afterAutospacing on MID body para) ===")
        ya = ys(word, build_generic('cas_body_auto.docx',
            para('上') + para('中後自動', 22, 'w:afterAutospacing="1"') + para('下')))
        yn = ys(word, build_generic('cas_body_norm.docx',
            para('上') + para('中普通') + para('下')))
        gA = ya[2][0]-ya[1][0]; gN = yn[2][0]-yn[1][0]
        print("  body  MID->BOT  auto=%.2f norm=%.2f  afterAuto=%.2f" % (gA, gN, gA-gN))

        # ---- (2) CELL value: autospacing on an INTERIOR cell para ----
        print("\n=== (2) CELL interior para (TOP / MID-auto / BOT all in one cell) ===")
        # afterAutospacing on MID
        cab = ys(word, build_generic('cas_cell_after.docx',
            cell(para('Ｔ') + para('中後自動', 22, 'w:afterAutospacing="1"') + para('Ｂ'))))
        cnb = ys(word, build_generic('cas_cell_norm.docx',
            cell(para('Ｔ') + para('中普通') + para('Ｂ'))))
        # cell paras: index 0=T,1=MID,2=B, then index3 = trailing body para after table
        gA = cab[2][0]-cab[1][0]; gN = cnb[2][0]-cnb[1][0]
        print("  cell  MID->BOT  after_auto=%.2f norm=%.2f  afterAuto=%.2f" % (gA, gN, gA-gN))
        # beforeAutospacing on MID
        cbb = ys(word, build_generic('cas_cell_before.docx',
            cell(para('Ｔ') + para('中前自動', 22, 'w:beforeAutospacing="1"') + para('Ｂ'))))
        gB = cbb[1][0]-cbb[0][0]; gNb = cnb[1][0]-cnb[0][0]
        print("  cell  TOP->MID  before_auto=%.2f norm=%.2f  beforeAuto=%.2f" % (gB, gNb, gB-gNb))

        # ---- (3) CELL EDGE suppression ----
        print("\n=== (3) CELL edge: MID is FIRST / LAST cell para ===")
        # MID = first cell para (before-auto at cell top)
        ef = ys(word, build_generic('cas_cell_firstb.docx',
            cell(para('先頭前自動', 22, 'w:beforeAutospacing="1"') + para('Ｂ'))))
        efn = ys(word, build_generic('cas_cell_firstn.docx',
            cell(para('先頭普通') + para('Ｂ'))))
        # For first-para: compare its Y to the table-top (no prior para). Use gap to next.
        # Better signal: distance from the cell paragraph[0] start to its own first line is
        # internal; instead compare cell para0 Y vs the same with normal first para.
        print("  first cell para Y  before_auto=%.2f  norm=%.2f  delta=%.2f"
              % (ef[0][0], efn[0][0], ef[0][0]-efn[0][0]))
        # MID = last cell para (after-auto at cell bottom): measure cell height via the
        # body para that follows the table.
        el = ys(word, build_generic('cas_cell_lasta.docx',
            cell(para('Ｔ') + para('末尾後自動', 22, 'w:afterAutospacing="1"')) + para('後続本文')))
        eln = ys(word, build_generic('cas_cell_lastn.docx',
            cell(para('Ｔ') + para('末尾普通')) + para('後続本文')))
        # the trailing body para is the last in each (index -1)
        print("  trailing-body Y after cell  after_auto=%.2f  norm=%.2f  delta=%.2f"
              % (el[-1][0], eln[-1][0], el[-1][0]-eln[-1][0]))

        # ---- (4) collapse: adjacent autospacing cell paras ----
        print("\n=== (4) CELL collapse A(after-auto)->B(before-auto) ===")
        col = ys(word, build_generic('cas_cell_collapse.docx',
            cell(para('Ｔ')
                 + para('Ａ後自動', 22, 'w:afterAutospacing="1"')
                 + para('Ｂ前自動', 22, 'w:beforeAutospacing="1"')
                 + para('Ｂ末'))))
        for i in range(len(col)-1):
            print("  %-8s -> %-8s gap=%.2f" % (col[i][1], col[i+1][1], col[i+1][0]-col[i][0]))

        # ---- (5) size sweep in cell ----
        print("\n=== (5) CELL afterAutospacing vs font size ===")
        print("  %-6s %-9s %-9s %-9s" % ("sz", "auto", "norm", "afterAuto"))
        for szhalf in (16, 21, 22, 24, 28, 32):
            fs = szhalf/2.0
            a = ys(word, build_generic('cas_sw_a_%d.docx'%szhalf,
                cell(para('Ｔ') + para('中', szhalf, 'w:afterAutospacing="1"') + para('Ｂ'))))
            n = ys(word, build_generic('cas_sw_n_%d.docx'%szhalf,
                cell(para('Ｔ') + para('中', szhalf) + para('Ｂ'))))
            gA = a[2][0]-a[1][0]; gN = n[2][0]-n[1][0]
            print("  %-6.1f %-9.2f %-9.2f %-9.2f" % (fs, gA, gN, gA-gN))
    finally:
        word.Quit()

if __name__ == '__main__':
    main()
