# -*- coding: utf-8 -*-
"""b35 per-cell render-truth breakdown. Input = rendertruth_match_pdf output (has oy, cell
(row,col), fs, rdy). Per-fs median calibration (removes ascent/baseline confound). Then group
by Oxi cell, cluster cell glyphs into lines by oy, and report per-cell-line median rdy2 with
line-index-in-cell. Aggregates: rdy2 by line-index-in-cell (first line vs continuation), by
cell row index (accumulation down the table). cp932-safe (ASCII out)."""
import json, sys, statistics, collections

m = json.load(open(sys.argv[1], encoding='utf-8'))['matched']
good = [g for g in m if abs(g['rdx']) < 5]
# per-fs median calib
byfs = collections.defaultdict(list)
for g in good:
    byfs[g['fs']].append(g['rdy'])
fsmed = {s: statistics.median(v) for s, v in byfs.items()}
for g in good:
    g['rdy2'] = g['rdy'] - fsmed[g['fs']]

# group by cell (row,col); None cell = body text (group as ('body',))
cells = collections.defaultdict(list)
for g in good:
    c = tuple(g['cell']) if g.get('cell') else (None, None)
    cells[c].append(g)

rows = []
for c, gs in cells.items():
    gs.sort(key=lambda g: g['oy'])
    # cluster into lines by oy
    lines = []
    for g in gs:
        if lines and abs(g['oy'] - lines[-1]['oy']) < 5:
            lines[-1]['gs'].append(g)
        else:
            lines.append({'oy': g['oy'], 'gs': [g]})
    for li, L in enumerate(lines):
        med = statistics.median(x['rdy2'] for x in L['gs'])
        rows.append({'cell': c, 'line_in_cell': li, 'n_lines': len(lines),
                     'oy': round(L['oy'], 0), 'n': len(L['gs']),
                     'fs': L['gs'][0]['fs'], 'rdy2': round(med, 2)})

# 1) by line-index-in-cell
print("=== rdy2 by LINE-INDEX-IN-CELL (0=first line; +=Oxi too low) ===")
byli = collections.defaultdict(list)
for r in rows:
    byli[r['line_in_cell']].append(r['rdy2'])
for li in sorted(byli):
    v = byli[li]
    print("  line %d in cell: n=%3d  median rdy2=%+5.2f  mean=%+5.2f" % (
        li, len(v), statistics.median(v), statistics.mean(v)))

# 2) by cell ROW index (table accumulation)
print("\n=== rdy2 by CELL ROW index (None=body) ===")
byrow = collections.defaultdict(list)
for r in rows:
    byrow[r['cell'][0]].append(r['rdy2'])
for rk in sorted(byrow, key=lambda x: (x is None, x)):
    v = byrow[rk]
    print("  row %-5s: n_lines=%3d  median rdy2=%+5.2f" % (str(rk), len(v), statistics.median(v)))

# 3) first-line vs continuation overall
firsts = [r['rdy2'] for r in rows if r['line_in_cell'] == 0 and r['n_lines'] > 1]
conts = [r['rdy2'] for r in rows if r['line_in_cell'] > 0]
singles = [r['rdy2'] for r in rows if r['n_lines'] == 1]
print("\n=== multi-line cells: first-line vs continuation ===")
if firsts:
    print("  first-line (of multi-line cells): n=%d median=%+.2f" % (len(firsts), statistics.median(firsts)))
if conts:
    print("  continuation lines:               n=%d median=%+.2f" % (len(conts), statistics.median(conts)))
if singles:
    print("  single-line cells:                 n=%d median=%+.2f" % (len(singles), statistics.median(singles)))
