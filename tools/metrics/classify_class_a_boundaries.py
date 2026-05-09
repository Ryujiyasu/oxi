"""Class A characterization (Day 32 Week 1):
For each monotone-positive doc, identify boundary paragraphs that flip
break decision between SOFT=0 and SOFT=N pt.

Output: per-doc list of:
- Total paragraphs
- Earliest divergent paragraph (= boundary trigger)
- # divergent paragraphs (= cascade extent)
- Boundary y position
- Pattern (transit page-break, line wrap shift, etc.)
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX_DIR = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')


def find_docx(doc_id: str) -> str | None:
    for f in os.listdir(DOCX_DIR):
        if f.startswith(doc_id) and f.endswith('.docx'):
            return os.path.join(DOCX_DIR, f)
    return None


def render(docx: str, margin: float) -> dict:
    env = os.environ.copy()
    env['OXI_SOFT_MARGIN_PT'] = str(margin)
    label = os.path.splitext(os.path.basename(docx))[0]
    out_layout = os.path.join(r'C:\tmp', f'{label}_sm{margin}_layout.json')
    cmd = [RENDERER, docx, os.path.join(r'C:\tmp', f'{label}_sm{margin}'), f'--dump-layout={out_layout}']
    r = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=180)
    if r.returncode != 0:
        return {}
    try:
        with open(out_layout, encoding='utf-8') as f:
            layout = json.load(f)
    except Exception:
        return {}
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
                by_para[pi] = {'first_pg': pg, 'first_y': el.get('y', 0), 'sample': el.get('text', '')[:30]}
            elif (pg, el.get('y', 0)) < (by_para[pi]['first_pg'], by_para[pi]['first_y']):
                by_para[pi]['first_pg'] = pg
                by_para[pi]['first_y'] = el.get('y', 0)
    return by_para


def compare(a: dict, b: dict) -> list[dict]:
    """Find paragraphs where position differs."""
    diff = []
    for pi in sorted(set(a.keys()) & set(b.keys())):
        ap = a[pi]
        bp = b[pi]
        if ap['first_pg'] != bp['first_pg'] or abs(ap['first_y'] - bp['first_y']) > 0.5:
            diff.append({
                'pi': pi,
                'a_pg': ap['first_pg'], 'a_y': round(ap['first_y'], 2),
                'b_pg': bp['first_pg'], 'b_y': round(bp['first_y'], 2),
                'sample': ap['sample'],
            })
    return diff


# 18 monotone-positive docs from drift_profile
DOC_IDS = [
    'e3c545fac7a7', 'a1d6e4efa2e7', '6514f214e482', 'cb8be715d839',
    'de6e32b5960b', 'd77a58485f16', 'd4d126dfe1d9', '191cb5254cb2',
    'bd90b00ab7a7', 'e201249db062', 'db9ca18368cd', '9a8e8ddab85b',
    '29dc6e8943fe', '6295e189a801', 'e8caed453f48', '15f9755cbccc',
    '8efcd416dfb8', 'a5ccbe425525',
]


def main():
    SOFT = 6.0  # value where bd90b00 transitions PASS
    print(f'Class A characterization: SOFT=0 vs SOFT={SOFT}pt')
    print()
    print(f'{"doc_id":<32} {"#paras":>6} {"diff":>6} {"first":>6} {"a_pg":>5} {"b_pg":>5} {"first sample"}')
    summary = []
    for doc_id in DOC_IDS:
        docx = find_docx(doc_id)
        if not docx:
            print(f'{doc_id:<32} NOT FOUND')
            continue
        a = render(docx, 0.0)
        b = render(docx, SOFT)
        if not a or not b:
            print(f'{doc_id:<32} RENDER FAILED')
            continue
        diff = compare(a, b)
        if diff:
            first = diff[0]
            print(f'{doc_id:<32} {len(a):>6} {len(diff):>6} {first["pi"]:>6} {first["a_pg"]:>5} {first["b_pg"]:>5} {first["sample"]!r}')
        else:
            print(f'{doc_id:<32} {len(a):>6} {0:>6} (no divergence)')
        summary.append({'doc_id': doc_id, 'n_paras': len(a), 'n_diff': len(diff), 'first_diff': diff[0] if diff else None})
    out_path = os.path.join(REPO, 'pipeline_data', 'class_a_boundaries.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump({'soft_margin': SOFT, 'docs': summary}, f, ensure_ascii=False, indent=2)
    print(f'\nWrote {out_path}')


if __name__ == '__main__':
    main()
