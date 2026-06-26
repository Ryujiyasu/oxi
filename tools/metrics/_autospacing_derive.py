# -*- coding: utf-8 -*-
"""Derive Word's beforeAutospacing/afterAutospacing rule (Ra, 2026-06-26).

Builds minimal repros (TOP / autospaced para / BOTTOM) and measures per-paragraph
Y via Information(6) with the R30 collapsed-start fix. Derives:
  (1) the auto value as a function of font size,
  (2) the COLLAPSE behavior between adjacent autospacing paragraphs (gate-safety),
  (3) suppression at page top / vs explicit-after neighbor.
"""
import os, sys, io
sys.path.insert(0, 'tools/metrics')
import win32com.client
from mixedh_lineplace import build_generic
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')

FONT = 'ＭＳ 明朝'

def rpr(sz=22, extra=''):
    return ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/>%s'
            % (FONT, FONT, FONT, sz, extra))

def para(text, sz=22, spacing_attrs='', ppr_extra=''):
    sp = ('<w:spacing %s/>' % spacing_attrs) if spacing_attrs else ''
    return ('<w:p><w:pPr>%s%s<w:rPr>%s</w:rPr></w:pPr>'
            '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (sp, ppr_extra, rpr(sz), rpr(sz), text))

def ys(word, docx):
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

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        # (1) size sweep — afterAutospacing only on the MID para
        print("=== (1) afterAutospacing value vs size ===")
        print("%-6s %-9s %-9s %-9s %-9s" % ("sz", "gap_auto", "gap_norm", "afterAuto", "ratio/fs"))
        for szhalf in (21, 22, 24, 28, 32):
            fs = szhalf / 2.0
            auto = build_generic('as_auto_%d.docx' % szhalf,
                para('上', 22) + para('自動後', szhalf, 'w:afterAutospacing="1"') + para('下', 22))
            norm = build_generic('as_norm_%d.docx' % szhalf,
                para('上', 22) + para('普通中', szhalf) + para('下', 22))
            ya = ys(word, auto); yn = ys(word, norm)
            gap_a = ya[2][0] - ya[1][0]   # MID -> BOTTOM
            gap_n = yn[2][0] - yn[1][0]
            aa = gap_a - gap_n
            print("%-6.1f %-9.2f %-9.2f %-9.2f %-9.3f" % (fs, gap_a, gap_n, aa, aa/fs))

        # (2) beforeAutospacing value vs size
        print("\n=== (2) beforeAutospacing value vs size ===")
        print("%-6s %-9s %-9s %-9s" % ("sz", "gap_auto", "gap_norm", "beforeAuto"))
        for szhalf in (21, 22, 24, 28):
            fs = szhalf / 2.0
            auto = build_generic('bs_auto_%d.docx' % szhalf,
                para('上', 22) + para('自動前', szhalf, 'w:beforeAutospacing="1"') + para('下', 22))
            norm = build_generic('bs_norm_%d.docx' % szhalf,
                para('上', 22) + para('普通中', szhalf) + para('下', 22))
            ya = ys(word, auto); yn = ys(word, norm)
            gap_a = ya[1][0] - ya[0][0]   # TOP -> MID
            gap_n = yn[1][0] - yn[0][0]
            print("%-6.1f %-9.2f %-9.2f %-9.2f" % (fs, gap_a, gap_n, gap_a - gap_n))

        # (3) COLLAPSE between two adjacent autospacing paragraphs (same default style)
        print("\n=== (3) collapse: A(after-auto) -> B(before-auto), both 11pt ===")
        d = build_generic('as_collapse.docx',
            para('上', 22)
            + para('A自動後', 22, 'w:afterAutospacing="1"')
            + para('B自動前', 22, 'w:beforeAutospacing="1"')
            + para('下', 22))
        y = ys(word, d)
        for i in range(len(y)-1):
            print("  %-8s -> %-8s gap=%.2f" % (y[i][1], y[i+1][1], y[i+1][0]-y[i][0]))
        # control: both normal
        dn = build_generic('as_collapse_norm.docx',
            para('上',22)+para('A普通',22)+para('B普通',22)+para('下',22))
        yn = ys(word, dn)
        print("  control normal A->B gap=%.2f (= line height)" % (yn[2][0]-yn[1][0]))

        # (4) two paras BOTH before+after auto (the kojin case), consecutive, same style
        print("\n=== (4) two consecutive paras both before+after auto (HTML-margin collapse?) ===")
        d = build_generic('as_both.docx',
            para('上', 22)
            + para('X両自動', 22, 'w:beforeAutospacing="1" w:afterAutospacing="1"')
            + para('Y両自動', 22, 'w:beforeAutospacing="1" w:afterAutospacing="1"')
            + para('下', 22))
        y = ys(word, d)
        for i in range(len(y)-1):
            print("  %-8s -> %-8s gap=%.2f" % (y[i][1], y[i+1][1], y[i+1][0]-y[i][0]))

        # (5) autospacing vs explicit after on neighbor (max collapse?)
        print("\n=== (5) A(after-auto) -> B(before=240=12pt explicit) ===")
        d = build_generic('as_vs_explicit.docx',
            para('上',22)
            + para('A自動後',22,'w:afterAutospacing="1"')
            + para('B明示前',22,'w:before="240"')
            + para('下',22))
        y = ys(word, d)
        for i in range(len(y)-1):
            print("  %-8s -> %-8s gap=%.2f" % (y[i][1], y[i+1][1], y[i+1][0]-y[i][0]))
    finally:
        word.Quit()

if __name__ == '__main__':
    main()
