"""Measure Word's actual line height for CI_* minimal repros.

For each docx:
1. Open with Word COM, get y of paragraph 1 inside cell, y of paragraph 2.
2. Word row gap = y_p2 - y_p1 = actual line height for that font/snap combo.
3. Compare to Oxi's predicted lh values from each formula candidate:
   - word_line_height_table_cell (current estimate path)
   - word_line_height_standard (current actual path when no GDI table)
   - GDI height × 83/64 (current actual path with GDI table)
   - GDI height × 1.0 (no inflate hypothesis)

Output: pipeline_data/cjk_inflate_sweep.json
"""
from __future__ import annotations
import os, sys, json, glob
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'cjk_inflate')
OUT = os.path.join(REPO, 'pipeline_data', 'cjk_inflate_sweep.json')


def measure(docx_path: str) -> dict:
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    out = {}
    try:
        n = d.Paragraphs.Count
        # Find first 2 paragraphs containing 'テスト'
        ys = []
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            t = (p.Range.Text or '')
            if 'テスト' in t:
                rng = p.Range
                cr = d.Range(rng.Start, rng.Start)
                y = cr.Information(6)
                ys.append((i, y, t.strip()[:20]))
                if len(ys) >= 2:
                    break
        if len(ys) >= 2:
            out['p1_i'] = ys[0][0]
            out['p1_y_pt'] = ys[0][1]
            out['p1_text'] = ys[0][2]
            out['p2_i'] = ys[1][0]
            out['p2_y_pt'] = ys[1][1]
            out['p2_text'] = ys[1][2]
            out['word_row_gap_pt'] = ys[1][1] - ys[0][1]
        else:
            out['error'] = f'only {len(ys)} テスト paragraphs found'
    finally:
        d.Close(SaveChanges=False)
        word.Quit()
    return out


def main():
    docs = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
    results = []
    for fp in docs:
        label = os.path.basename(fp).replace('.docx', '')
        # parse fs and snap from label CI_fs{N}_snap{S}
        parts = label.split('_')
        fs_str = parts[1].replace('fs', '').replace('p', '.')
        fs = float(fs_str)
        snap = int(parts[2].replace('snap', ''))
        m = measure(fp)
        m['label'] = label
        m['fs_pt'] = fs
        m['snap'] = snap
        results.append(m)
        gap = m.get('word_row_gap_pt', None)
        gap_s = f'{gap:.3f}' if gap is not None else 'ERR'
        print(f'{label:30} fs={fs:>4}  snap={snap}  word_gap={gap_s}pt')
    json.dump(results, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'\nSaved: {OUT}')

    # Pretty pivot
    print()
    print(f'{"font_size":>10} {"snap=0":>10} {"snap=1":>10} {"diff":>8}')
    by_fs = {}
    for r in results:
        if 'word_row_gap_pt' in r:
            by_fs.setdefault(r['fs_pt'], {})[r['snap']] = r['word_row_gap_pt']
    for fs in sorted(by_fs.keys()):
        s0 = by_fs[fs].get(0, '?')
        s1 = by_fs[fs].get(1, '?')
        diff = (s1 - s0) if isinstance(s0, float) and isinstance(s1, float) else '?'
        s0s = f'{s0:.3f}' if isinstance(s0, float) else str(s0)
        s1s = f'{s1:.3f}' if isinstance(s1, float) else str(s1)
        ds = f'{diff:+.3f}' if isinstance(diff, float) else str(diff)
        print(f'{fs:>10} {s0s:>10} {s1s:>10} {ds:>8}')

if __name__ == '__main__':
    main()
