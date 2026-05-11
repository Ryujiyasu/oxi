"""Day 33 part 13 — measure Oxi cell line height post-fix.

Loads cjk_inflate_sweep_v2.json (Word measurements) and runs Oxi GDI renderer
with --dump-layout to extract cell paragraph y values. Compares Oxi vs Word.
"""
from __future__ import annotations
import os, sys, json, glob, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'cjk_inflate_v2')
WORD_DATA = os.path.join(REPO, 'pipeline_data', 'cjk_inflate_sweep_v2.json')
OUT = os.path.join(REPO, 'pipeline_data', 'cjk_inflate_oxi_v2.json')
GDI_BIN = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')


def measure_oxi(docx_path: str, label: str) -> dict:
    layout_json = os.path.join(r'C:\tmp', f'{label}_oxi_v2.json')
    out_prefix = os.path.join(r'C:\tmp', label + '_oxi_v2')
    os.makedirs(r'C:\tmp', exist_ok=True)
    res = subprocess.run([
        GDI_BIN, docx_path, out_prefix, '96',
        f'--dump-layout={layout_json}',
    ], capture_output=True, text=True)
    if res.returncode != 0:
        return {'error': res.stderr[:300]}
    d = json.load(open(layout_json, encoding='utf-8'))
    found = []
    for p in d['pages']:
        for el in p['elements']:
            if el['type'] == 'text' and 'テスト' in el.get('text', ''):
                found.append({'y': el['y'], 'text': el['text'][:20], 'fs': el.get('font_size'), 'h': el.get('h')})
                if len(found) >= 2:
                    break
        if len(found) >= 2:
            break
    if len(found) >= 2:
        return {
            'p1_y_pt': found[0]['y'],
            'p1_text': found[0]['text'],
            'p2_y_pt': found[1]['y'],
            'p2_text': found[1]['text'],
            'oxi_row_gap_pt': round(found[1]['y'] - found[0]['y'], 4),
        }
    return {'error': f'only {len(found)} paragraphs found'}


def main():
    with open(WORD_DATA, 'r', encoding='utf-8') as f:
        word_data = {r['label']: r for r in json.load(f)}
    docs = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
    results = []
    for fp in docs:
        label = os.path.basename(fp).replace('.docx', '')
        wd = word_data.get(label, {})
        word_gap = wd.get('word_row_gap_pt')
        oxi = measure_oxi(fp, label)
        oxi['label'] = label
        oxi['font'] = wd.get('font')
        oxi['fs_pt'] = wd.get('fs_pt')
        oxi['word_row_gap_pt'] = word_gap
        if 'oxi_row_gap_pt' in oxi and word_gap is not None:
            oxi['diff_pt'] = oxi['oxi_row_gap_pt'] - word_gap
        results.append(oxi)
        og = oxi.get('oxi_row_gap_pt')
        og_s = f'{og:.4f}' if og is not None else f'ERR'
        wg_s = f'{word_gap:.4f}' if word_gap is not None else '?'
        diff = oxi.get('diff_pt')
        diff_s = f'{diff:+.4f}' if diff is not None else '?'
        print(f'{label:32}  oxi={og_s}  word={wg_s}  diff={diff_s}')
    json.dump(results, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'\nSaved: {OUT}')

    # Pivot: diff per font/fs
    print()
    print(f'{"font":>6}{"fs=8":>10}{"fs=9":>10}{"fs=10":>10}{"fs=10.5":>10}{"fs=11":>10}{"fs=12":>10}{"fs=14":>10}')
    by_font = {}
    for r in results:
        if 'diff_pt' in r and r.get('font'):
            by_font.setdefault(r['font'], {})[r['fs_pt']] = r['diff_pt']
    for font in ('MSM', 'MSG', 'YuM', 'YuG'):
        row = [f'{font:>6}']
        for fs in (8, 9, 10, 10.5, 11, 12, 14):
            v = by_font.get(font, {}).get(fs, '?')
            row.append(f'{v:>+10.3f}' if isinstance(v, float) else f'{v:>10}')
        print(''.join(row))

if __name__ == '__main__':
    main()
