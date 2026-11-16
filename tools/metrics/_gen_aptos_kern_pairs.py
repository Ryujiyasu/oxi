# -*- coding: utf-8 -*-
"""Extract Aptos' GPOS 'kern' pair adjustments (ASCII + curly quotes) for Oxi.

en/forms__006d1fbe (Aptos 12, docDefaults w:kern 2): Word sets «have» with v->e
5.26pt where hmtx gives 5.43 -- the GPOS pair v,e. Oxi's latin_kern table had no
Aptos entry, so every kerned Aptos line ran ~0.3% wide.

Reads the Office cloud-font cache (the font files stay on the machine) and writes
crates/oxidocs-core/src/font/data/aptos_kern_pairs.json keyed by full face name.

    python tools/metrics/_gen_aptos_kern_pairs.py [--write]
"""
import glob
import json
import os
import sys

from fontTools.ttLib import TTFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _gen_meiryo_kern_pairs import pairs_of  # noqa: E402

CACHE = os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft/FontCache/4/CloudFonts')

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    table = {}
    for fam in ('Aptos', 'Aptos Display', 'Aptos Narrow'):
        for path in sorted(glob.glob(os.path.join(CACHE, fam, '*.ttf'))):
            f = TTFont(path)
            full = f['name'].getDebugName(4)
            p = pairs_of(f)
            print(full, f['head'].unitsPerEm, len(p), {k: p.get(k) for k in ('118,101', '84,111', '65,86', '87,97')})
            table[full] = p
    if '--write' in sys.argv:
        out = 'crates/oxidocs-core/src/font/data/aptos_kern_pairs.json'
        with open(out, 'w', encoding='utf-8') as fh:
            json.dump(table, fh, ensure_ascii=False, separators=(',', ':'), sort_keys=True)
        print('written', out, len(table))
