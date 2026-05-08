"""Measure Oxi's actual line height for CI_* minimal repros via dump-layout.

Reads the layout JSON dump from oxi-gdi-renderer for each variant, finds
paragraph fragments matching 'テスト１' and 'テスト２', computes y diff =
Oxi's actual rendered line height.

Compares to Word measurements from cjk_inflate_sweep.json.

Output: pipeline_data/cjk_inflate_oxi.json + side-by-side comparison table.
"""
from __future__ import annotations
import os, sys, json, glob, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'cjk_inflate')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'
WORD_JSON = os.path.join(REPO, 'pipeline_data', 'cjk_inflate_sweep.json')
OUT = os.path.join(REPO, 'pipeline_data', 'cjk_inflate_oxi.json')


def measure_oxi(docx_path: str, label: str) -> dict:
    layout_json = os.path.join(TMP, f'{label}_oxi.json')
    out_prefix = os.path.join(TMP, label + '_oxi')
    res = subprocess.run([
        RENDERER, docx_path, out_prefix, '96',
        f'--dump-layout={layout_json}',
    ], capture_output=True, text=True)
    if res.returncode != 0:
        return {'error': res.stderr[:300]}
    d = json.load(open(layout_json, encoding='utf-8'))
    out = {}
    # Find first 2 elements containing 'テスト'
    found = []
    for p in d['pages']:
        for el in p['elements']:
            if el['type'] == 'text' and 'テスト' in el.get('text', ''):
                found.append({'y': el['y'], 'text': el['text'][:20], 'fs': el['font_size'], 'h': el['h']})
                if len(found) >= 2:
                    break
        if len(found) >= 2:
            break
    if len(found) >= 2:
        out['p1'] = found[0]
        out['p2'] = found[1]
        out['oxi_row_gap_pt'] = round(found[1]['y'] - found[0]['y'], 4)
    return out


def main():
    # Load Word measurements
    word_data = json.load(open(WORD_JSON, encoding='utf-8'))
    word_by_label = {r['label']: r for r in word_data}

    docs = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
    results = []
    for fp in docs:
        label = os.path.basename(fp).replace('.docx', '')
        m = measure_oxi(fp, label)
        m['label'] = label
        # Parse fs and snap
        parts = label.split('_')
        fs = float(parts[1].replace('fs', '').replace('p', '.'))
        snap = int(parts[2].replace('snap', ''))
        m['fs_pt'] = fs
        m['snap'] = snap
        wd = word_by_label.get(label, {})
        m['word_row_gap_pt'] = wd.get('word_row_gap_pt')
        if 'oxi_row_gap_pt' in m and m['word_row_gap_pt'] is not None:
            m['diff_pt'] = round(m['oxi_row_gap_pt'] - m['word_row_gap_pt'], 4)
        results.append(m)
    json.dump(results, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)

    # Print comparison table
    print(f'{"label":35} {"fs":>5} {"snap":>5} {"word":>8} {"oxi":>8} {"diff":>8}')
    print('-' * 75)
    for r in sorted(results, key=lambda x: (x['fs_pt'], x['snap'])):
        wg = r.get('word_row_gap_pt')
        og = r.get('oxi_row_gap_pt')
        df = r.get('diff_pt')
        wg_s = f'{wg:.3f}' if wg is not None else '?'
        og_s = f'{og:.3f}' if og is not None else '?'
        df_s = f'{df:+.3f}' if df is not None else '?'
        print(f'{r["label"]:35} {r["fs_pt"]:>5} {r["snap"]:>5} {wg_s:>8} {og_s:>8} {df_s:>8}')

    print(f'\nSaved: {OUT}')

if __name__ == '__main__':
    main()
