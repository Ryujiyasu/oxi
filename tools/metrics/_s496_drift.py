# -*- coding: utf-8 -*-
"""S496 font-fair per-line drift profile (render-truth, PDF route).
Word glyphs: word_pdf_glyphs.py output (baseline y in pt). Oxi glyphs: dwrite
--dump-glyphs with the EXACT 'baseline' field (S494b font-fair, no K guess).
Content-match via difflib, calibrate by median, print per-line mean rdy/rdx vs line_y.
cp932-safe: ASCII out, results written to a file. rdy>0 = Oxi BELOW Word (too low)."""
import json, sys, difflib, statistics


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


def profile(wpath, opath, pidx, outf):
    W = json.load(open(wpath, encoding='utf-8'))['pages'][pidx]['glyphs']
    O = json.load(open(opath, encoding='utf-8'))['pages'][pidx]['glyphs']
    wg = [{'char': g['char'], 'x': g['x'], 'by': g['y'], 'fs': g.get('fs', 0)} for g in W if g['char'].strip()]
    og = [{'char': g['char'], 'x': g['x'], 'by': g['baseline'], 'fs': g['font_size']} for g in O if g['char'].strip()]
    wseq = [g for L in cluster(wg) for g in L['gs']]
    oseq = [g for L in cluster(og) for g in L['gs']]
    sm = difflib.SequenceMatcher(None, [g['char'] for g in wseq],
                                 [g['char'] for g in oseq], autojunk=False)
    matched = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            for d in range(i2 - i1):
                matched.append((wseq[i1 + d], oseq[j1 + d]))
    L = []
    if not matched:
        L.append('NO MATCH p%d' % pidx)
        outf.write('\n'.join(L) + '\n'); return
    cal_y = statistics.median(o['by'] - w['by'] for w, o in matched)
    cal_x = statistics.median(o['x'] - w['x'] for w, o in matched)
    lines = {}
    for w, o in matched:
        rdy = (o['by'] - w['by']) - cal_y
        rdx = (o['x'] - w['x']) - cal_x
        key = round(w['by'], 0)
        lines.setdefault(key, {'rdy': [], 'rdx': [], 'fs': []})
        lines[key]['rdy'].append(rdy); lines[key]['rdx'].append(rdx); lines[key]['fs'].append(w['fs'])
    L.append('page %d  match_ratio %.3f  matched %d/%d  cal_y %.2f cal_x %.2f  (rdy>0=Oxi LOW)'
             % (pidx, sm.ratio(), len(matched), len(wseq), cal_y, cal_x))
    L.append(' line_y   n   fs   rdy_mean rdy_std  rdx_mean')
    ys = sorted(lines)
    for y in ys:
        r = lines[y]
        rym = statistics.mean(r['rdy']); rys = statistics.pstdev(r['rdy']) if len(r['rdy']) > 1 else 0
        rxm = statistics.mean(r['rdx']); fs = statistics.median(r['fs'])
        L.append('  %6.1f %3d %4.1f  %+6.2f  %5.2f  %+6.2f' % (y, len(r['rdy']), fs, rym, rys, rxm))
    allr = [(y, statistics.mean(lines[y]['rdy'])) for y in ys]
    if len(allr) > 2:
        first3 = statistics.mean(r for _, r in allr[:3])
        last3 = statistics.mean(r for _, r in allr[-3:])
        span = allr[-1][0] - allr[0][0]
        L.append('TOP-3 rdy %+.2f -> BOT-3 rdy %+.2f  (delta %+.2f over %.0fpt = %+.3f pt/100pt)'
                 % (first3, last3, last3 - first3, span, (last3 - first3) / span * 100 if span else 0))
    outf.write('\n'.join(L) + '\n')


def main():
    wpath, opath = sys.argv[1], sys.argv[2]
    pages = sys.argv[3]  # e.g. "3" or "0,3,5" or "all"
    out = sys.argv[4] if len(sys.argv) > 4 else 'c:/tmp/_s496_drift_out.txt'
    n = len(json.load(open(wpath, encoding='utf-8'))['pages'])
    if pages == 'all':
        idxs = list(range(n))
    else:
        idxs = [int(x) for x in pages.split(',')]
    with open(out, 'w', encoding='utf-8') as f:
        for pi in idxs:
            profile(wpath, opath, pi, f)
            f.write('\n')
    print('wrote', out)


if __name__ == '__main__':
    main()
