"""Measure PB_DRIFT_CUMUL — full y trajectory of N=30 identical paragraphs.

Tracks per-paragraph advance (y[i+1] - y[i]) for both Word and Oxi, and
the cumulative drift after k paragraphs. Reveals where Word's cumulative
ceil diverges from Oxi's per-paragraph round.
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, '..', '..', '..'))
GDI_EXE = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

VARIANTS = [
    ("DR_CU_01", "ＭＳ 明朝", 10.5, "auto"),
    ("DR_CU_02", "ＭＳ 明朝", 11.0, "auto"),
    ("DR_CU_03", "ＭＳ 明朝", 11.5, "auto"),
    ("DR_CU_04", "ＭＳ 明朝", 12.0, "auto"),
    ("DR_CU_05", "ＭＳ 明朝", 14.0, "auto"),
    ("DR_CU_06", "ＭＳ 明朝", 10.5, "multiple1.15"),
    ("DR_CU_07", "ＭＳ 明朝", 10.5, "multiple1.5"),
    ("DR_CU_08", "ＭＳ 明朝", 10.5, "exact14"),
    ("DR_CU_09", "ＭＳ 明朝", 11.5, "exact16"),
    ("DR_CU_10", "Yu Mincho", 10.5, "auto"),
]


def measure_word(docx_path):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False; word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        ys = []
        n = d.Paragraphs.Count
        for i in range(1, n + 1):
            rng = d.Paragraphs(i).Range
            cr = d.Range(rng.Start, rng.Start)
            ys.append((int(cr.Information(3)), round(cr.Information(6), 3)))
        return ys
    finally:
        d.Close(False); word.Quit()


def measure_oxi(docx_path):
    layout_path = os.path.join(HERE, '_tmp_layout.json')
    out_prefix = os.path.join(HERE, '_tmp_oxi')
    cmd = [GDI_EXE, docx_path, out_prefix, '150', f'--dump-layout={layout_path}']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0: return []
    with open(layout_path, encoding='utf-8') as f:
        layout = json.load(f)
    out = []
    for page_idx, page in enumerate(layout.get('pages', [])):
        # Pre-build per-pi min y on this page
        seen = {}
        for el in page.get('elements', []):
            if el.get('type') != 'text': continue
            pi = el.get('para_idx')
            if pi is None: continue
            y = el.get('y', 0)
            if pi not in seen or y < seen[pi]:
                seen[pi] = y
        for pi in sorted(seen):
            out.append((page_idx + 1, seen[pi]))
    return out


def main():
    summary = []
    for vid, font, fs, lh in VARIANTS:
        path = os.path.join(HERE, f'{vid}.docx')
        if not os.path.exists(path): continue
        print(f'\n=== {vid}: {font} {fs}pt lh={lh} ===')
        wd = measure_word(path)
        ox = measure_oxi(path)
        n = min(len(wd), len(ox))
        if n < 5:
            print(f'  too few paragraphs: word={len(wd)} oxi={len(ox)}'); continue
        # Per-paragraph cumulative drift (oxi_y - word_y) on same page
        rows = []
        for i in range(n):
            wp, wy = wd[i]
            op, oy = ox[i]
            if wp != op:
                rows.append({'i': i+1, 'wp': wp, 'wy': wy, 'op': op, 'oy': oy, 'dy': None, 'page_diff': True})
            else:
                rows.append({'i': i+1, 'wp': wp, 'wy': wy, 'op': op, 'oy': oy, 'dy': round(oy-wy, 3), 'page_diff': False})
        # Per-paragraph advance
        for i in range(1, len(rows)):
            r_prev = rows[i-1]; r_now = rows[i]
            if r_now['wp'] == r_prev['wp']:
                r_now['w_adv'] = round(r_now['wy'] - r_prev['wy'], 3)
            if r_now['op'] == r_prev['op']:
                r_now['o_adv'] = round(r_now['oy'] - r_prev['oy'], 3)
        # Print first 10 + last 5
        print(f'  {"i":>3} {"wp":>3} {"wy":>7} {"op":>3} {"oy":>7} {"dy":>7} {"w_adv":>7} {"o_adv":>7}')
        idxs = list(range(min(10, len(rows)))) + ([len(rows)-1] if len(rows) > 10 else [])
        for j in idxs:
            r = rows[j]
            d = f'{r["dy"]:+7.3f}' if r['dy'] is not None else '   ~~~~'
            wa = f'{r.get("w_adv","-"):>7.3f}' if 'w_adv' in r else '   -   '
            oa = f'{r.get("o_adv","-"):>7.3f}' if 'o_adv' in r else '   -   '
            print(f'  {r["i"]:>3} {r["wp"]:>3} {r["wy"]:>7.2f} {r["op"]:>3} {r["oy"]:>7.2f} {d} {wa} {oa}')

        # Aggregate
        same_page = [r for r in rows if not r['page_diff']]
        if len(same_page) >= 2:
            dys = [r['dy'] for r in same_page]
            init_dy = dys[0]; final_dy = dys[-1]
            print(f'  init_dy={init_dy:+.3f}  final_dy={final_dy:+.3f}  range_drift={final_dy-init_dy:+.3f}')

        summary.append({
            'variant': vid, 'font': font, 'fs': fs, 'lh': lh,
            'n_word': len(wd), 'n_oxi': len(ox),
            'rows': rows,
        })

    out = os.path.join(HERE, 'measurements.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f'\nSaved: {out}')

    # Cross-variant cumulative drift summary
    print('\n=== Cumulative drift summary (final_dy - init_dy after N=30) ===')
    print(f'{"variant":<10} {"font":<12} {"fs":>5} {"lh":<14} {"init":>7} {"final":>7} {"range":>7}')
    for s in summary:
        rows = s['rows']
        same_page = [r for r in rows if not r['page_diff']]
        if len(same_page) < 2: continue
        init = same_page[0]['dy']; final = same_page[-1]['dy']
        print(f'  {s["variant"]:<10} {s["font"]:<12} {s["fs"]:>5.1f} {s["lh"]:<14} {init:>+7.3f} {final:>+7.3f} {final-init:>+7.3f}')


if __name__ == '__main__':
    main()
