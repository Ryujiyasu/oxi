"""Measure Word vs Oxi wrap line counts on DW_V113-V115 repros.

For each variant:
  - Word COM: count lines via Range loops, compute paragraph height in pt
  - Oxi: --dump-layout, count text elements per paragraph, compute total height
  - Compare → identify wrap miscount (Oxi over by N lines = +18N pt drift)
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'db9ca_wrap')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'

VARIANTS = [
    'DW_V113_para18_only',
    'DW_V114_para17_only',
    'DW_V115_combined_17_18_19',
]


def measure_word(docx_path: str) -> dict:
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    out: dict = {}
    paras = []
    try:
        ps = d.PageSetup
        out['pgH'] = round(ps.PageHeight, 2)
        out['top_margin'] = round(ps.TopMargin, 2)
        out['bottom_margin'] = round(ps.BottomMargin, 2)
        # For each paragraph, get start y and end y
        n_paras = d.Paragraphs.Count
        for i in range(1, n_paras + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr_start = d.Range(r.Start, r.Start)
            cr_end = d.Range(r.End - 1, r.End - 1)
            text = (r.Text or '').strip()
            if not text:
                continue
            paras.append({
                'i': i,
                'text': text[:60],
                'len': len(text),
                'start_y': round(cr_start.Information(6), 2),
                'end_y': round(cr_end.Information(6), 2),
                'start_page': int(cr_start.Information(3)),
                'end_page': int(cr_end.Information(3)),
            })
        out['paras'] = paras
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
        return {'error': r.stderr[-300:]}
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    pages = layout.get('pages', [])
    # Group text elements by para_idx
    paras = {}
    for page_idx, page in enumerate(pages):
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            pi = el.get('para_idx')
            if pi is None:
                continue
            if pi not in paras:
                paras[pi] = {
                    'para_idx': pi,
                    'lines': [],
                    'first_y': el['y'],
                    'last_y': el['y'],
                    'text_concat': '',
                }
            paras[pi]['lines'].append({'y': el['y'], 'h': el['h'], 'text': el.get('text', '')[:30], 'page': page_idx + 1})
            paras[pi]['last_y'] = max(paras[pi]['last_y'], el['y'])
            paras[pi]['first_y'] = min(paras[pi]['first_y'], el['y'])
            paras[pi]['text_concat'] += el.get('text', '')
    # Count distinct y-values per paragraph (= number of lines)
    out_paras = []
    for pi, p in sorted(paras.items()):
        unique_ys = sorted(set(round(line['y'], 1) for line in p['lines']))
        out_paras.append({
            'para_idx': pi,
            'n_lines': len(unique_ys),
            'first_y': round(p['first_y'], 2),
            'last_y': round(p['last_y'], 2),
            'text_len': len(p['text_concat']),
        })
    return {'paras': out_paras}


def main():
    results = []
    for label in VARIANTS:
        docx = os.path.join(REPRO_DIR, f'{label}.docx')
        if not os.path.exists(docx):
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
        # Print Word paragraphs
        print('  Word paragraphs:')
        for p in w['paras']:
            n_lines = max(1, round((p['end_y'] - p['start_y']) / 18.0) + 1) if p['end_page'] == p['start_page'] else '?'
            print(f'    i={p["i"]} pg={p["start_page"]}->{p["end_page"]} y={p["start_y"]}->{p["end_y"]} '
                  f'(span={p["end_y"]-p["start_y"]:.1f}pt ~ {n_lines} lines) len={p["len"]}')
        # Print Oxi paragraphs
        print('  Oxi paragraphs:')
        for p in o.get('paras', []):
            print(f'    para_idx={p["para_idx"]} n_lines={p["n_lines"]} '
                  f'y_span={p["first_y"]}->{p["last_y"]} (={p["last_y"]-p["first_y"]:.1f}pt) text_len={p["text_len"]}')
        results.append({'label': label, 'word': w, 'oxi': o})
        print()
    out_path = os.path.join(REPO, 'pipeline_data', 'db9ca_wrap_results.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'Wrote {out_path}')


if __name__ == '__main__':
    main()
