# -*- coding: utf-8 -*-
"""Oxi readout for tests/fixtures/vertwrap: per-band strip counts and x ranges,
from the S1411 column trace (OXI_DBG_VCOLS), next to Word's _word.json.

Word rows are (page, x, y) of each paragraph's collapsed start; Oxi rows are
the placed column's left edge x and start y. Word's Information(5) reports
the column's LEFT edge too (control: 523.5 = 541.3 - 18), so the numbers
compare directly.
"""
import os, sys, json, subprocess
from pathlib import Path
from collections import OrderedDict
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
GDI = os.environ.get('OXI_GDI_EXE') or 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe'
OUT = Path('tests/fixtures/vertwrap')
word = json.loads((OUT / '_word.json').read_text(encoding='utf-8'))
only = set(sys.argv[1:])
tmp = Path(os.environ.get('TEMP', '.')) / 'vertwrap_oxi'
tmp.mkdir(exist_ok=True)
for name in word:
    if only and name not in only:
        continue
    env = dict(os.environ, OXI_DBG_VCOLS='1')
    r = subprocess.run([GDI, str(OUT / f'{name}.docx'), str(tmp / name)], env=env, capture_output=True, timeout=300)
    rows = []
    npages = 0
    err = r.stderr.decode('utf-8', 'replace').splitlines()
    last = max((i for i, l in enumerate(err) if l.startswith('[VPASS]')), default=-1)
    for line in err[last + 1:]:
        if line.startswith('[VCOLS]'):
            kv = dict(t.split('=') for t in line.split()[1:])
            rows.append((int(kv['page']) + 1, float(kv['x']), float(kv['y'])))
        if line.startswith('Parsed'):
            npages = int(line.split()[1])
    bands = OrderedDict()
    for (pg, x, y) in rows:
        bands.setdefault((pg, round(y, 2)), []).append(x)
    o = '; '.join(f'p{pg} y{y}: n={len(xs)} x {max(xs)}..{min(xs)}' for (pg, y), xs in bands.items())
    wb = OrderedDict()
    for (pg, x, y) in word[name]['rows']:
        wb.setdefault((pg, y), []).append(x)
    w = '; '.join(f'p{pg} y{y}: n={len(xs)} x {max(xs)}..{min(xs)}' for (pg, y), xs in wb.items())
    key = lambda d: [(len(xs), round(max(xs) / 1.5), round(min(xs) / 1.5), round(y / 1.5)) for (pg, y), xs in d.items()]
    same = key(wb) == key(bands)
    print(f'{name:16s} {"OK " if same else "DIFF"} pages W{word[name]["pages"]}/O{npages}\n   W {w}\n   O {o}')
