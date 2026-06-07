# -*- coding: utf-8 -*-
"""S502 Word border render-truth: extract table column boundaries (vertical line segments)
from the Word PDF via PyMuPDF get_drawings(). The Word PDF has VECTOR borders (unlike the
glyph dump), so this gives Word's actual column x-positions per page. Compare to Oxi's
layout-dump vertical borders to localize column-width errors. cp932-safe.
Usage: python _s502_word_borders.py <docx_or_rtpdf> <page_index>"""
import sys, os, glob
import fitz

def main():
    arg = sys.argv[1]
    pidx = int(sys.argv[2]) if len(sys.argv) > 2 else 0
    if arg.endswith('.pdf'):
        pdf = arg
    else:
        base = os.path.splitext(os.path.abspath(arg))[0]
        pdf = base + '_rt.pdf'
        if not os.path.exists(pdf):
            cands = glob.glob(base + '*_rt.pdf') or glob.glob(os.path.dirname(base) + '/*_rt.pdf')
            pdf = cands[0] if cands else pdf
    if not os.path.exists(pdf):
        print('PDF not found:', pdf); return
    d = fitz.open(pdf)
    page = d[pidx]
    H = page.rect.height
    # collect vertical line segments (x ~const, dy large) from drawings
    vlines = {}  # x(rounded) -> total dy
    for dr in page.get_drawings():
        for item in dr['items']:
            if item[0] == 'l':  # line: ('l', p1, p2)
                p1, p2 = item[1], item[2]
                if abs(p1.x - p2.x) < 0.6 and abs(p1.y - p2.y) > 8:
                    x = round((p1.x + p2.x) / 2, 1)
                    vlines[x] = vlines.get(x, 0) + abs(p1.y - p2.y)
            elif item[0] == 're':  # rect: edges count as v-lines at x0,x1
                r = item[1]
                if r.height > 8:
                    for x in (round(r.x0, 1), round(r.x1, 1)):
                        vlines[x] = vlines.get(x, 0) + r.height
    # keep x with substantial vertical extent (real column borders)
    cols = sorted(x for x, dy in vlines.items() if dy > 20)
    print('Word PDF p%d vertical borders (col boundaries): %s' % (pidx, cols))
    print('page width %.1f' % page.rect.width)

if __name__ == '__main__':
    main()
