# -*- coding: utf-8 -*-
"""S494 grind — per-LINE vertical drift profile (render-truth, PDF route).
Word glyphs: word_pdf_glyphs.py output (baseline y in pt). Oxi glyphs: dwrite
--dump-glyphs (top + fs*K baseline, K matches the MuPDF gate render). Content-match
via difflib, calibrate cal_y = median(oxi_bl - word_bl), then cluster Word glyphs into
lines and print per-line mean rdy vs line_y. If rdy grows ~linearly with y the page has
a per-line PITCH error (systematic vertical accumulation). cp932-safe, ASCII out."""
import json, sys, difflib, statistics

K = 0.859  # ascender ratio used by oxi_via_mupdf render (baseline = top + fs*K)


def cluster(glyphs, tol=4.0):
    gs = sorted(glyphs, key=lambda g: (round(g['by'], 1), g['x']))
    lines = []
    for g in gs:
        if lines and abs(g['by'] - lines[-1]['by']) < tol:
            lines[-1]['gs'].append(g)
        else:
            lines.append({'by': g['by'], 'gs': [g]})
    for L in lines:
        L['gs'].sort(key=lambda g: g['x'])
        L['by'] = statistics.median(g['by'] for g in L['gs'])
    return lines


def main():
    wpath, opath, pidx = sys.argv[1], sys.argv[2], int(sys.argv[3])
    W = json.load(open(wpath, encoding='utf-8'))['pages'][pidx]['glyphs']
    O = json.load(open(opath, encoding='utf-8'))['pages'][pidx]['glyphs']
    wg = [{'char': g['char'], 'x': g['x'], 'by': g['y']} for g in W if g['char'].strip()]
    og = [{'char': g['char'], 'x': g['x'], 'by': g['top'] + g['font_size'] * K,
           'fs': g['font_size']} for g in O if g['char'].strip()]

    wseq = [g for L in cluster(wg) for g in L['gs']]
    oseq = [g for L in cluster(og) for g in L['gs']]
    sm = difflib.SequenceMatcher(None, [g['char'] for g in wseq],
                                 [g['char'] for g in oseq], autojunk=False)
    matched = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            for d in range(i2 - i1):
                matched.append((wseq[i1 + d], oseq[j1 + d]))
    if not matched:
        print('NO MATCH'); return
    cal_y = statistics.median(o['by'] - w['by'] for w, o in matched)
    cal_x = statistics.median(o['x'] - w['x'] for w, o in matched)
    # per-line: group matched by Word baseline
    lines = {}
    for w, o in matched:
        rdy = (o['by'] - w['by']) - cal_y
        rdx = (o['x'] - w['x']) - cal_x
        lines.setdefault(round(w['by'], 0), []).append((rdy, rdx))
    print('page %d  match_ratio %.3f  matched %d  cal_y %.2f cal_x %.2f  (rdy>0=Oxi LOW)'
          % (pidx, sm.ratio(), len(matched), cal_y, cal_x))
    print(' line_y   n   rdy_mean  rdx_mean')
    ys = sorted(lines)
    for y in ys:
        rs = lines[y]
        rym = statistics.mean(r[0] for r in rs)
        rxm = statistics.mean(r[1] for r in rs)
        print('  %6.1f %3d   %+6.2f   %+6.2f' % (y, len(rs), rym, rxm))
    # slope estimate: rdy vs line index
    allr = [(y, statistics.mean(r[0] for r in lines[y])) for y in ys]
    if len(allr) > 2:
        n = len(allr)
        first3 = statistics.mean(r for _, r in allr[:3])
        last3 = statistics.mean(r for _, r in allr[-3:])
        span = allr[-1][0] - allr[0][0]
        print('TOP-3 lines rdy_mean %+.2f -> BOTTOM-3 rdy_mean %+.2f  (delta %+.2f over %.0fpt)'
              % (first3, last3, last3 - first3, span))


if __name__ == '__main__':
    main()
