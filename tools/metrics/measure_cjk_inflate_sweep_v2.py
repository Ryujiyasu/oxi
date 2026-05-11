"""Day 33 part 13 — COM-measure cjk_inflate_v2 repros.

For each docx, measure y of paragraph 1 (text 'テスト１') and paragraph 2
('テスト２') in the cell. row_gap = y2 - y1 = actual line height for
(font, fs) at snap=0 in cell.

Output: pipeline_data/cjk_inflate_sweep_v2.json
"""
from __future__ import annotations
import os, sys, json, glob
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'cjk_inflate_v2')
OUT = os.path.join(REPO, 'pipeline_data', 'cjk_inflate_sweep_v2.json')


def measure(word, docx_path: str) -> dict:
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    out = {}
    try:
        n = d.Paragraphs.Count
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
    return out


def main():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    docs = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
    results = []
    try:
        for fp in docs:
            label = os.path.basename(fp).replace('.docx', '')
            # parse: CJK2_{font}_fs{N}_snap0
            parts = label.split('_')
            font_short = parts[1]
            fs_str = parts[2].replace('fs', '').replace('p', '.')
            fs = float(fs_str)
            m = measure(word, fp)
            m['label'] = label
            m['font'] = font_short
            m['fs_pt'] = fs
            results.append(m)
            gap = m.get('word_row_gap_pt', None)
            gap_s = f'{gap:.4f}' if gap is not None else f'ERR: {m.get("error","?")}'
            print(f'{label:32} font={font_short:4}  fs={fs:>4}  word_gap={gap_s}pt')
    finally:
        word.Quit()
    json.dump(results, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'\nSaved: {OUT}')

    # Pivot
    print()
    print(f'{"font":>6}{"fs=8":>10}{"fs=9":>10}{"fs=10":>10}{"fs=10.5":>10}{"fs=11":>10}{"fs=12":>10}{"fs=14":>10}')
    by_font = {}
    for r in results:
        if 'word_row_gap_pt' in r:
            by_font.setdefault(r['font'], {})[r['fs_pt']] = r['word_row_gap_pt']
    for font in ('MSM', 'MSG', 'YuM', 'YuG'):
        row = [f'{font:>6}']
        for fs in (8, 9, 10, 10.5, 11, 12, 14):
            v = by_font.get(font, {}).get(fs, '?')
            row.append(f'{v:>10.3f}' if isinstance(v, float) else f'{v:>10}')
        print(''.join(row))

if __name__ == '__main__':
    main()
