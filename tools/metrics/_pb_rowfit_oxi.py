# -*- coding: utf-8 -*-
"""Oxi readout for tests/fixtures/rowfit: page/y of row 1 line 1, row 1 line 2, row 2."""
import os, sys, json, subprocess
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
GDI = os.environ.get('OXI_GDI_EXE') or 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe'
OUT = Path('tests/fixtures/rowfit'); tmp = Path(os.environ.get('TEMP', '.')) / 'rowfit_oxi'; tmp.mkdir(exist_ok=True)
names = sys.argv[1:] or sorted(p.stem for p in OUT.glob('*.docx'))
for name in names:
    dump = tmp / f'{name}.json'
    subprocess.run([GDI, str(OUT / f'{name}.docx'), str(tmp / name), f'--dump-layout={dump}'], capture_output=True, timeout=300)
    d = json.load(open(dump, encoding='utf-8'))
    l1 = l2 = r2 = None
    for pi, pg in enumerate(d['pages'], 1):
        for e in pg['elements']:
            if e.get('type') != 'text' or e.get('cell_row_idx') is None: continue
            if e['cell_row_idx'] == 0 and e.get('cell_para_idx') == 0 and l1 is None: l1 = (pi, round(e['y'], 2))
            if e['cell_row_idx'] == 0 and e.get('cell_para_idx') == 1 and l2 is None: l2 = (pi, round(e['y'], 2))
            if e['cell_row_idx'] == 1 and r2 is None: r2 = (pi, round(e['y'], 2))
    print(f'{name:9s} row1 line1 {l1} line2 {l2} | row2 {r2}')
