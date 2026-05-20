"""S136: Measure TR_V200-V203 minimal repros to falsify Hypothesis H1.

H1: `is_first_block_est` subtraction at mod.rs:6302-6317 subtracts
space_before from row_height; if H1 is correct:
  V200 (before-only)  → drift ≈ -4.35pt/row × n_tables
  V201 (after-only)   → drift ≈ +1.0pt/row (= V101 baseline; sb subtraction doesn't fire)
  V202 (before + trH) → drift ≈ +1.0pt/row (trHeight overrides the underestimate)
  V203 (both, sanity) → drift ≈ -2.74pt/row (matches V100 if our V100 template is equivalent)

Output: pipeline_data/tokumei_sb_isolation_results.json
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'tokumei_slow_drift')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'

VARIANTS = [
    'TR_V200_before_only',
    'TR_V201_after_only',
    'TR_V202_before_with_trheight',
    'TR_V203_both_sanity',
]


def measure_word(docx_path: str) -> dict:
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    out: dict = {}
    rows = []
    try:
        ps = d.PageSetup
        out['pgH'] = round(ps.PageHeight, 3)
        out['top_margin'] = round(ps.TopMargin, 3)
        out['bottom_margin'] = round(ps.BottomMargin, 3)
        n = d.Paragraphs.Count
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            txt = (p.Range.Text or '').strip()
            if txt.startswith('当該公的機関の名称'):
                rng = p.Range
                cr = d.Range(rng.Start, rng.Start)
                rows.append({
                    'i': i,
                    'text': txt[:30],
                    'page': int(cr.Information(3)),
                    'y': round(cr.Information(6), 3),
                    'x': round(cr.Information(5), 3),
                    'in_table': bool(cr.Information(12)),
                })
        out['rows'] = rows
    finally:
        d.Close(False)
        word.Quit()
    return out


def measure_oxi(docx_path: str) -> dict:
    label = os.path.splitext(os.path.basename(docx_path))[0]
    out_prefix = os.path.join(TMP, f'{label}')
    out_layout = os.path.join(TMP, f'{label}_layout.json')
    cmd = [RENDERER, docx_path, out_prefix, f'--dump-layout={out_layout}']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        return {'error': r.stderr[-500:]}
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    rows = []
    pages = layout.get('pages', [])
    for page_idx, page in enumerate(pages):
        elements = page.get('elements', [])
        # one paragraph per row, dedupe by para_idx + cell_row_idx
        seen = set()
        for el in elements:
            if el.get('type') != 'text':
                continue
            txt = (el.get('text') or '').strip()
            if not txt.startswith('当該公的機関の名称'):
                continue
            key = (el.get('para_idx'), el.get('cell_row_idx'), el.get('cell_col_idx'))
            if key in seen:
                continue
            seen.add(key)
            rows.append({
                'page': page.get('page', page_idx + 1),
                'text': txt[:30],
                'y': round(el.get('y', 0), 3),
                'x': round(el.get('x', 0), 3),
            })
    return {'rows': rows}


def analyze(label: str, w: dict, o: dict) -> dict:
    wr = w.get('rows', [])
    or_ = o.get('rows', [])
    n = min(len(wr), len(or_))
    pairs = []
    for i in range(n):
        w_row, o_row = wr[i], or_[i]
        page_h = w.get('pgH', 841.95)
        w_y_abs = (w_row['page'] - 1) * page_h + w_row['y']
        o_y_abs = (o_row['page'] - 1) * page_h + o_row['y']
        dy_abs = o_y_abs - w_y_abs
        pairs.append({
            'idx': i,
            'w_page': w_row['page'], 'o_page': o_row['page'],
            'w_y': w_row['y'], 'o_y': o_row['y'],
            'dy_abs': round(dy_abs, 3),
        })
    if pairs:
        cum_drift = pairs[-1]['dy_abs'] - pairs[0]['dy_abs']
        per_row = cum_drift / max(1, n - 1)
    else:
        cum_drift = 0
        per_row = 0
    return {
        'label': label,
        'n_w': len(wr), 'n_o': len(or_),
        'cum_drift_pt': round(cum_drift, 3),
        'per_row_drift_pt': round(per_row, 4),
        'pairs': pairs,
    }


def main():
    results = []
    for label in VARIANTS:
        docx = os.path.join(REPRO_DIR, f'{label}.docx')
        if not os.path.exists(docx):
            print(f'SKIP {label}: not found')
            continue
        print(f'=== {label} ===')
        try:
            w = measure_word(docx)
        except Exception as e:
            print(f'  Word ERROR: {e}')
            continue
        try:
            o = measure_oxi(docx)
        except Exception as e:
            print(f'  Oxi ERROR: {e}')
            continue
        if 'error' in o:
            print(f'  Oxi render error: {o["error"]}')
            continue
        a = analyze(label, w, o)
        print(f'  n_word={a["n_w"]} n_oxi={a["n_o"]} '
              f'cum_drift={a["cum_drift_pt"]:+.3f}pt '
              f'per_row={a["per_row_drift_pt"]:+.4f}pt/row')
        for i in (0, 5, 10, 15, 20, 25, len(a['pairs']) - 1):
            if 0 <= i < len(a['pairs']):
                p = a['pairs'][i]
                print(f'    i={p["idx"]:>2} w_pg={p["w_page"]} o_pg={p["o_page"]} '
                      f'w_y={p["w_y"]:>6.2f} o_y={p["o_y"]:>6.2f} '
                      f'dy={p["dy_abs"]:>+6.3f}')
        results.append(a)
        print()
    out_path = os.path.join(REPO, 'pipeline_data', 'tokumei_sb_isolation_results.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'Wrote {out_path}')


if __name__ == '__main__':
    main()
