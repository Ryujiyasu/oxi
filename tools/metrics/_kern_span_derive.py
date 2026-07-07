# -*- coding: utf-8 -*-
"""Per-SPAN rigorous Latin break-metric derivation: for every PDF text span
(its OWN font + size), compare Word's per-char advances against em and
em+kern. Aggregates per (family, size). Skips space-adjacent advances
(justify stretch) and cross-span pairs.

Usage: python tools/metrics/_kern_span_derive.py <word_pdf> [min_size max_size]
"""
import sys, os, json
sys.stdout.reconfigure(encoding='utf-8')
import fitz
from fontTools.ttLib import TTFont

FONT_FILES = {
    'TimesNewRomanPSMT': ('Times New Roman', r'C:\Windows\Fonts\times.ttf'),
    'TimesNewRomanPS-BoldMT': ('Times New Roman Bold', r'C:\Windows\Fonts\timesbd.ttf'),
    'Century': ('Century', r'C:\Windows\Fonts\CENTURY.TTF'),
    'ArialMT': ('Arial', r'C:\Windows\Fonts\arial.ttf'),
    'Arial-BoldMT': ('Arial Bold', r'C:\Windows\Fonts\arialbd.ttf'),
    'Calibri': ('Calibri', r'C:\Windows\Fonts\calibri.ttf'),
    'Calibri-Bold': ('Calibri Bold', r'C:\Windows\Fonts\calibrib.ttf'),
    'Cambria': ('Cambria', r'C:\Windows\Fonts\cambria.ttc'),
}

_kern_cache = {}
def kern_table(path):
    if path not in _kern_cache:
        f = TTFont(path, fontNumber=0) if path.endswith('.ttc') else TTFont(path)
        k = {}
        if 'kern' in f:
            for st in f['kern'].kernTables:
                k.update(st.kernTable)
        _kern_cache[path] = (f.getBestCmap(), k, f['head'].unitsPerEm)
    return _kern_cache[path]

def main():
    pdf = sys.argv[1]
    fm = json.load(open('crates/oxidocs-core/src/font/data/font_metrics_compact.json', encoding='utf-8'))
    tables = {e['family']: (e['widths'], e['units_per_em']) for e in fm}
    d = fitz.open(pdf)
    agg = {}
    for pno in range(len(d)):
        for b in d[pno].get_text('rawdict')['blocks']:
            for l in b.get('lines', []):
                for s in l['spans']:
                    fname = s['font']
                    if fname not in FONT_FILES:
                        continue
                    fam, path = FONT_FILES[fname]
                    if fam not in tables:
                        continue
                    w, upm = tables[fam]
                    cmap, kern, kupm = kern_table(path)
                    # PDF size is scaled (10.56 for 10.5); round to the half-point
                    fs = round(s['size'] * 2) / 2
                    # correct the PDF scale quirk (10.56 -> 10.5, 12.0 -> 12.0)
                    fs_eff = round(s['size'] / 1.0057 * 2) / 2 if abs(s['size'] - round(s['size'])) > 0.01 else s['size']
                    chars = s['chars']
                    for i in range(len(chars) - 1):
                        c = chars[i]['c']; nxt = chars[i+1]['c']
                        if c == ' ' or nxt == ' ':
                            continue
                        a = w.get(str(ord(c)))
                        if a is None:
                            continue
                        adv = chars[i+1]['bbox'][0] - chars[i]['bbox'][0]
                        if adv <= 0 or adv > fs_eff * 1.6:
                            continue  # line-wrap artifacts
                        em_w = a / upm * fs_eff
                        k = 0.0
                        ga, gb = cmap.get(ord(c)), cmap.get(ord(nxt))
                        if ga and gb:
                            ku = kern.get((ga, gb))
                            if ku:
                                k = ku / kupm * fs_eff
                        key = (fam, fs_eff)
                        e = agg.setdefault(key, [0, 0.0, 0.0, 0.0])
                        e[0] += 1; e[1] += adv; e[2] += em_w; e[3] += em_w + k
    print(f'{"family":<24} {"fs":>5} {"n":>6} {"Word":>9} {"em d":>8} {"em+kern d":>10}  per-char')
    for (fam, fs), (n, ws, es, ks) in sorted(agg.items()):
        print(f'{fam:<24} {fs:>5} {n:>6} {ws:>9.1f} {es-ws:>+8.2f} {ks-ws:>+10.2f}  {(ks-ws)/n:+.4f}')

if __name__ == '__main__':
    main()
