# -*- coding: utf-8 -*-
"""Per-line rdx profile from a rendertruth_match_glyph output. For each matched line (cluster by oy):
report n, rdx at first glyph, rdx at last glyph, mean rdx, in-line slope (rdx_last - rdx_first).
Distinguishes: justify-spread (rdx ramps within line), uniform per-line shift (flat nonzero),
accumulation. Also flags last char x vs Word last char x. cp932-safe (JSON in/out, ASCII)."""
import json, sys, statistics
m = json.load(open(sys.argv[1], encoding='utf-8'))['matched']
m.sort(key=lambda g: (round(g['oy'], 1), g['ox']))
lines = []
for g in m:
    if lines and abs(g['oy'] - lines[-1]['oy']) < 5:
        lines[-1]['gs'].append(g)
    else:
        lines.append({'oy': g['oy'], 'gs': [g]})
print("per-line rdx (Oxi-Word horizontal resid, calib-removed). slope=last-first; n=glyphs")
print("rampUP(+slope)=Oxi spreads more; rampDOWN(-)=Oxi packs tighter than Word's justify")
rows = []
for L in lines:
    gs = sorted(L['gs'], key=lambda g: g['ox'])
    rf, rl = gs[0]['rdx'], gs[-1]['rdx']
    mean = statistics.mean(g['rdx'] for g in gs)
    slope = rl - rf
    # absolute Word/Oxi line-end x (un-calibrated) to see who's wider
    wxe = gs[-1]['ox'] - gs[-1]['rdx']  # reconstruct: ox - rdx = wx + cal; but we want span — use raw ox/wx
    rows.append({'oy': round(L['oy'], 1), 'n': len(gs), 'rf': round(rf, 1), 'rl': round(rl, 1),
                 'mean': round(mean, 2), 'slope': round(slope, 1)})
for r in rows:
    bar = ''
    if r['slope'] <= -3: bar = '  <== Oxi packs tighter (ends '+str(r['rl'])+'pt left)'
    elif r['slope'] >= 3: bar = '  ==> Oxi spreads wider'
    print("  oy=%6.1f n=%2d rdx_first=%+5.1f rdx_last=%+5.1f mean=%+5.2f slope=%+5.1f%s" % (
        r['oy'], r['n'], r['rf'], r['rl'], r['mean'], r['slope'], bar))
slopes = [r['slope'] for r in rows]
means = [r['mean'] for r in rows]
print("\nSUMMARY: n_lines=%d  slope mean=%+.2f std=%.2f  | per-line mean: mean=%+.2f std=%.2f" % (
    len(rows), statistics.mean(slopes), statistics.pstdev(slopes), statistics.mean(means), statistics.pstdev(means)))
neg = [r for r in rows if r['slope'] <= -3]
print("lines with slope<=-3pt (Oxi packs >=3pt tighter by line-end): %d / %d" % (len(neg), len(rows)))
print("  these are likely FULL (justified) lines where Word spreads to right margin, Oxi doesn't")
