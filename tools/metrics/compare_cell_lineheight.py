# -*- coding: utf-8 -*-
"""S492v — compare Oxi vs Word cell line pitch on the gen_cell_lineheight repros.
Dump Oxi layout (GDI --dump-layout) per repro, extract the cell text line-Y pitch, compare
to the COM-measured Word pitch (cell_lineheight_word.json). Confirms Oxi over-spaces
snapToGrid=0 cell lines (uniform ~17.5/18.0) vs Word's small natural height. ASCII out."""
import os, glob, json, subprocess
from pathlib import Path
import numpy as np

ROOT = Path(os.path.abspath('.'))
GDI = str(ROOT / 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
REPRO = r'c:\tmp\cellrepro'
word = {r['name']: r for r in json.load(open(r'c:\tmp\cell_lineheight_word.json', encoding='utf-8'))}


def oxi_cell_pitch(docx):
    dpath = 'c:/tmp/_cl_dump.json'
    subprocess.run([GDI, docx, 'c:/tmp/_cl_out', '150', '--dump-layout=' + dpath],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if not os.path.exists(dpath):
        return None, None, None
    d = json.load(open(dpath, encoding='utf-8'))
    els = [e for e in d['pages'][0]['elements'] if e.get('type') == 'text']
    if not els:
        return None, None, None
    ys = sorted(set(round(e['y'] * 2) / 2 for e in els))
    merged = []
    for y in ys:
        if merged and y - merged[-1] < 3:
            continue
        merged.append(y)
    deltas = [round(merged[k] - merged[k - 1], 2) for k in range(1, len(merged))]
    sd = sorted(deltas)
    med = sd[len(sd) // 2] if sd else None
    return med, len(merged), sorted(set(deltas))


print("CELL line pitch: Word (COM, snapToGrid=0 cell) vs Oxi (GDI dump)  [grid linePitch=17.5]")
print("font     size   Word_pitch  Oxi_pitch   Oxi-Word   Oxi_deltas")
for path in sorted(glob.glob(os.path.join(REPRO, '*.docx'))):
    name = os.path.basename(path)
    w = word.get(name)
    wp = w['line_pitch_pt'] if w else None
    op, nl, od = oxi_cell_pitch(path)
    ft = 'Mincho' if 'Mincho' in name else 'Gothic'
    sz = w['size_pt'] if w else 0
    diff = ('%+.2f' % (op - wp)) if (op is not None and wp is not None) else '  -  '
    print("  %-7s %5.1f    %s       %s      %s    %s" % (
        ft, sz,
        ('%.2f' % wp) if wp is not None else ' - ',
        ('%.2f' % op) if op is not None else ' - ',
        diff, od))
