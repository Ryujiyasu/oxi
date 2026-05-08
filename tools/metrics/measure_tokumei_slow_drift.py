"""Measure TS_V100-V106 minimal repros: Word vs Oxi per-row y position
across 30 consecutive 1-cell tables. Goal: localize tokumei sub-family
slow accumulation drift (+0.05-0.10pt/cell observed in d4d126/de6e/6514/a1d6/191cb).

For each variant:
  - Word COM: paragraph i=1..30, get y position via Information(6) from
              collapsed-start range (R30 fix).
  - Oxi: --dump-layout=, find each row paragraph y from text element list.
  - Compute per-row delta and total drift over 30 rows.
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'tokumei_slow_drift')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'

VARIANTS = [
    'TS_V100_baseline',
    'TS_V101_no_before_after',
    'TS_V102_no_valign',
    'TS_V103_lineRule_auto',
    'TS_V104_no_flag',
    'TS_V105_line200',
    'TS_V106_line300',
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
        # Match paragraphs that start with "当該公的機関の名称"
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
    # Extract text elements with text starting with "当該公的機関の名称"
    rows = []
    pages = layout.get('pages', [])
    for page_idx, page in enumerate(pages):
        elements = page.get('elements', [])
        for el in elements:
            if el.get('type') != 'text':
                continue
            txt = (el.get('text') or '').strip()
            if txt.startswith('当該公的機関の名称'):
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
        # absolute y across pages
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
    # Cumulative drift: dy_abs[n-1] - dy_abs[0]
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
        # Show pairs at indices 0, 5, 10, 15, 20, 25, 29
        for i in (0, 5, 10, 15, 20, 25, len(a['pairs']) - 1):
            if 0 <= i < len(a['pairs']):
                p = a['pairs'][i]
                print(f'    i={p["idx"]:>2} w_pg={p["w_page"]} o_pg={p["o_page"]} '
                      f'w_y={p["w_y"]:>6.2f} o_y={p["o_y"]:>6.2f} '
                      f'dy={p["dy_abs"]:>+6.3f}')
        results.append(a)
        print()
    out_path = os.path.join(REPO, 'pipeline_data', 'tokumei_slow_drift_results.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'Wrote {out_path}')


if __name__ == '__main__':
    main()
