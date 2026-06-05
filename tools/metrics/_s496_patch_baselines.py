# -*- coding: utf-8 -*-
"""S496: patch ssim_baseline.json and per_glyph_baseline.json for the 3 affected
mode-14 docs using the ALREADY-MEASURED ON values from the gate output files
(methodology verified identical: OFF column == stored baseline). Loads each baseline,
updates only the 3 docs' per-page values, writes back. cp932-safe."""
import json, io, re, sys

DWRITE_OUT = sys.argv[1]   # bhmgv3dpc.output (dwrite ON values)  -> ssim_baseline.json
PERGLYPH_OUT = sys.argv[2] # brg0vglcf.output (per-glyph ON)      -> per_glyph_baseline.json
SSIM = 'pipeline_data/ssim_baseline.json'
PG = 'pipeline_data/per_glyph_baseline.json'
TARGETS = ['e3c545', '04b88e', '34140b']

LINE = re.compile(r'^(\S.*?)\s+p(\d+)\s+([\d.]+)\s+([\d.]+)\s+([+\-][\d.]+)\s+\(([\d.]+)\)')

def parse_on(path):
    """returns {doc_prefix: {page_str: on_value}}"""
    out = {}
    for ln in io.open(path, encoding='utf-8'):
        m = LINE.match(ln.strip())
        if not m:
            continue
        doc = m.group(1).strip(); pg = m.group(2); on = float(m.group(4))
        out.setdefault(doc, {})[pg] = on
    return out

def patch(basepath, onmap):
    base = json.load(io.open(basepath, encoding='utf-8'))
    changed = []
    for doc_disp, pages in onmap.items():
        # find full key by prefix
        pref = next((t for t in TARGETS if doc_disp.startswith(t)), None)
        if not pref:
            continue
        full = [k for k in base if k.startswith(pref)][0]
        oldm = sum(base[full].values()) / len(base[full])
        for pg, v in pages.items():
            base[full][pg] = v
        newm = sum(base[full].values()) / len(base[full])
        changed.append((full, oldm, newm))
    json.dump(base, io.open(basepath, 'w', encoding='utf-8'), indent=2, ensure_ascii=False)
    return changed

print('=== ssim_baseline.json ===')
for full, o, n in patch(SSIM, parse_on(DWRITE_OUT)):
    print('  %-40s %.4f -> %.4f  (%+.4f)' % (full[:40], o, n, n - o))
print('=== per_glyph_baseline.json ===')
for full, o, n in patch(PG, parse_on(PERGLYPH_OUT)):
    print('  %-40s %.4f -> %.4f  (%+.4f)' % (full[:40], o, n, n - o))
print('DONE')
