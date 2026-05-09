"""Per-paragraph trace of 0e7af1ae8f21 at SOFT=0 vs SOFT=0.5pt to find
the earliest divergent paragraph (cascade trigger).

Day 31 part 31: Long-term per-paragraph detailed trace foundation.
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx',
                    '0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')


def render_at_margin(margin: float, label: str) -> dict:
    """Render at given SOFT_MARGIN, return {para_idx: {first_y, page, n_lines}}."""
    env = os.environ.copy()
    env['OXI_SOFT_MARGIN_PT'] = str(margin)
    out_prefix = os.path.join(r'C:\tmp', f'0e7af_{label}')
    out_layout = os.path.join(r'C:\tmp', f'0e7af_{label}_layout.json')
    cmd = [RENDERER, DOCX, out_prefix, f'--dump-layout={out_layout}']
    r = subprocess.run(cmd, capture_output=True, text=True, env=env)
    if r.returncode != 0:
        print(f'render @ {margin}pt failed: {r.stderr[:300]}')
        return {}
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    by_para = {}
    for page in layout.get('pages', []):
        pg = page.get('page')
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            pi = el.get('para_idx')
            if pi is None:
                continue
            if pi not in by_para:
                by_para[pi] = {
                    'para_idx': pi,
                    'first_y': el.get('y'),
                    'first_page': pg,
                    'last_y': el.get('y'),
                    'last_page': pg,
                    'pages': set(),
                    'ys': set(),
                    'sample_text': el.get('text', ''),
                }
            d = by_para[pi]
            d['pages'].add(pg)
            d['ys'].add(round(el.get('y', 0), 1))
            if (pg, el.get('y')) < (d['first_page'], d['first_y']):
                d['first_y'] = el.get('y')
                d['first_page'] = pg
            if (pg, el.get('y')) > (d['last_page'], d['last_y']):
                d['last_y'] = el.get('y')
                d['last_page'] = pg
    return {pi: {
        'para_idx': pi,
        'first_y': round(v['first_y'], 2),
        'first_page': v['first_page'],
        'n_lines': len(v['ys']),
        'pages': sorted(v['pages']),
        'sample_text': v['sample_text'][:50],
    } for pi, v in by_para.items()}


def main():
    print('Rendering at SOFT=0pt...')
    a = render_at_margin(0.0, 'soft0')
    print(f'  {len(a)} paragraphs')
    print('Rendering at SOFT=0.5pt...')
    b = render_at_margin(0.5, 'soft05')
    print(f'  {len(b)} paragraphs')
    # Find earliest divergent paragraph
    common = sorted(set(a.keys()) & set(b.keys()))
    diverge = []
    for pi in common:
        ap = a[pi]
        bp = b[pi]
        if ap['first_page'] != bp['first_page'] or abs(ap['first_y'] - bp['first_y']) > 0.5:
            diverge.append({
                'pi': pi,
                'a_pg': ap['first_page'], 'a_y': ap['first_y'], 'a_lines': ap['n_lines'],
                'b_pg': bp['first_page'], 'b_y': bp['first_y'], 'b_lines': bp['n_lines'],
                'sample': ap['sample_text'],
            })
    print(f'\n{len(diverge)} divergent paragraphs')
    if diverge:
        print('\nFirst 10 divergent (sorted by para_idx):')
        for d in diverge[:10]:
            print(f'  pi={d["pi"]:>3}: SOFT=0 pg={d["a_pg"]}/y={d["a_y"]} lines={d["a_lines"]} | SOFT=0.5 pg={d["b_pg"]}/y={d["b_y"]} lines={d["b_lines"]} | {d["sample"]!r}')
        print(f'\nEarliest divergent paragraph: pi={diverge[0]["pi"]}')
    out_path = os.path.join(REPO, 'pipeline_data', '0e7af_cascade_trace.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump({'soft0': a, 'soft05': b, 'divergent': diverge}, f, ensure_ascii=False, indent=2, default=list)
    print(f'\nWrote {out_path}')


if __name__ == '__main__':
    main()
