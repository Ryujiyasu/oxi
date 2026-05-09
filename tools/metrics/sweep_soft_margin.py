"""Sweep OXI_SOFT_MARGIN_PT values and measure Phase 1 (n_pass, mean_score).

Day 31 part 28: find optimal SOFT_MARGIN value for Phase 1 +1 PASS.
"""
from __future__ import annotations
import os, sys, subprocess, json
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))


def run_with_margin(margin: float) -> dict:
    """Run pagination_oxi + pagination_diff with given SOFT_MARGIN."""
    env = os.environ.copy()
    env['OXI_SOFT_MARGIN_PT'] = str(margin)
    # Clear caches
    pag_oxi = os.path.join(REPO, 'pipeline_data', 'pagination_oxi')
    for f in os.listdir(pag_oxi):
        if f.endswith('.json'):
            os.remove(os.path.join(pag_oxi, f))
    # Run pagination_oxi
    r = subprocess.run(
        ['python', 'tools/metrics/measure_pagination_oxi.py'],
        cwd=REPO, env=env, capture_output=True, text=True, encoding='utf-8', errors='replace'
    )
    if r.returncode != 0:
        print(f'  measure_pagination_oxi failed: {r.stderr[-500:]}')
        return {}
    # Run pagination_diff
    r = subprocess.run(
        ['python', 'tools/metrics/pagination_diff.py'],
        cwd=REPO, capture_output=True, text=True, encoding='utf-8', errors='replace'
    )
    if r.returncode != 0:
        print(f'  pagination_diff failed: {r.stderr[-500:]}')
        return {}
    # Parse summary
    with open(os.path.join(REPO, 'pipeline_data', 'pagination_diff', '_summary.json'), encoding='utf-8') as f:
        s = json.load(f)
    return {
        'n_pass': s['n_pass'],
        'n_total': s['n_total'],
        'mean_score': s['mean_score'],
    }


def main():
    values = [0.0, 3.0, 5.0, 7.0, 10.0, 15.0, 20.0]
    results = []
    for v in values:
        print(f'Testing SOFT_MARGIN={v:.1f}pt...')
        r = run_with_margin(v)
        if r:
            print(f'  → n_pass={r["n_pass"]}/{r["n_total"]} mean={r["mean_score"]:.4f}')
            results.append({'soft_margin': v, **r})
    print()
    print(f'{"margin":>7} {"pass":>6} {"mean":>8}')
    for r in results:
        print(f'{r["soft_margin"]:>7.1f} {r["n_pass"]}/{r["n_total"]:<3} {r["mean_score"]:>8.4f}')
    out_path = os.path.join(REPO, 'pipeline_data', 'soft_margin_sweep.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'\nWrote {out_path}')


if __name__ == '__main__':
    main()
