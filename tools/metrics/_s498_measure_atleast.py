# -*- coding: utf-8 -*-
"""Measure cell content placement for the atLeast repros: Word (PDF) vs Oxi (dump-glyphs).
For each variant report: content(あ) baseline Y, REF baseline Y, AFTER baseline Y, and
content-minus-REF (the cell-content offset). Reveals where each engine puts atLeast-cell
content. cp932-safe: matches by ascii REF/AFTER + the first non-ascii glyph for content."""
import os, sys, json, subprocess, tempfile
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
REPRO = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'atleast_cellY')

def word_glyphs(docx):
    out = tempfile.mktemp(suffix='.json')
    subprocess.run([sys.executable, os.path.join(ROOT, 'tools', 'metrics', 'word_pdf_glyphs.py'), docx, out],
                   capture_output=True, timeout=120)
    return json.load(open(out, encoding='utf-8'))['pages'][0]['glyphs']

def oxi_glyphs(docx):
    gj = tempfile.mktemp(suffix='.json')
    subprocess.run([DW, docx, tempfile.mktemp(), '150', '--dump-glyphs=' + gj], capture_output=True, timeout=120)
    return json.load(open(gj, encoding='utf-8'))['pages'][0]['glyphs']

def find(glyphs, pred, ykey):
    for g in glyphs:
        if pred(g['char']):
            return g[ykey]
    return None

def main():
    print('variant      | Word: REF   content  c-REF | Oxi: REF   content  c-REF | dC(O-W)')
    for a in ['none', '20', '30', '40', '50']:
        dx = os.path.join(REPRO, 'atleast_%s.docx' % a)
        if not os.path.exists(dx):
            continue
        wg = word_glyphs(dx); og = oxi_glyphs(dx)
        # Word baseline = y; Oxi baseline = 'baseline' field
        wref = find(wg, lambda c: c == 'R', 'y')
        wcon = find(wg, lambda c: ord(c) > 0x3000, 'y')
        oref = find(og, lambda c: c == 'R', 'baseline')
        ocon = find(og, lambda c: ord(c) > 0x3000, 'baseline')
        if None in (wref, wcon, oref, ocon):
            print('%-12s | MISSING wref=%s wcon=%s oref=%s ocon=%s' % (a, wref, wcon, oref, ocon))
            continue
        wcr = wcon - wref; ocr = ocon - oref
        print('%-12s | %7.2f %7.2f %+6.2f | %7.2f %7.2f %+6.2f | %+.2f'
              % (a, wref, wcon, wcr, oref, ocon, ocr, ocr - wcr))

if __name__ == '__main__':
    main()
