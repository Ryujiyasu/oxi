"""Per-element drift profile tool — extends element_iou_diff data with
cumulative-drift analysis to identify compensation patterns across PASS
and FAIL docs.

For each doc with element_iou_diff data:
  1. Read matched elements (Word vs Oxi y/h positions)
  2. Compute per-element absolute drift (= oxi_y_abs - word_y_abs)
  3. Compute cumulative drift trajectory + per-paragraph delta
  4. Classify into:
       - "matching" (cum_drift stays bounded ±2pt)
       - "monotone-positive" (Oxi consistently lower → over-pump cascade)
       - "monotone-negative" (Oxi consistently higher → under-pump)
       - "compensating" (cum_drift returns to ±2pt after deviation)
       - "outlier" (large jumps in delta)

Output:
  pipeline_data/drift_profile/_summary.json — aggregate stats
  pipeline_data/drift_profile/<doc_id>.json — per-doc profile
"""
from __future__ import annotations
import json
import os
import sys
import glob
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
IN_DIR = os.path.join(REPO, 'pipeline_data', 'element_iou_diff')
OUT_DIR = os.path.join(REPO, 'pipeline_data', 'drift_profile')

PAGE_HEIGHT = 841.95  # A4 portrait pt


def classify_trajectory(deltas: list[float]) -> str:
    if not deltas:
        return 'empty'
    n = len(deltas)
    abs_max = max(abs(d) for d in deltas)
    if abs_max < 2.0:
        return 'matching'
    # Trend slope from linear regression on indices
    avg = sum(deltas) / n
    # Spread vs trend
    if avg > 2.0 and all(d > -2.0 for d in deltas):
        return 'monotone-positive' if deltas[-1] > deltas[0] else 'plateau-positive'
    if avg < -2.0 and all(d < 2.0 for d in deltas):
        return 'monotone-negative' if deltas[-1] < deltas[0] else 'plateau-negative'
    # Look for compensation: end value smaller than max excursion
    if abs_max > 5.0 and abs(deltas[-1]) < abs_max * 0.5:
        return 'compensating'
    # Outlier check: largest single jump > 50% of total range
    if n >= 2:
        jumps = [abs(deltas[i] - deltas[i-1]) for i in range(1, n)]
        if jumps and max(jumps) > 30.0:
            return 'outlier-jumps'
    return 'mixed'


def analyze_doc(data: dict) -> dict:
    matched = [m for m in data.get('matches', []) if m.get('matched')]
    matched.sort(key=lambda m: (m['word_page'], m['word_y']))
    deltas = []
    for m in matched:
        word_y_abs = (m['word_page'] - 1) * PAGE_HEIGHT + m['word_y']
        oxi_y_abs = (m['oxi_page'] - 1) * PAGE_HEIGHT + m['oxi_y']
        deltas.append(oxi_y_abs - word_y_abs)

    trajectory = classify_trajectory(deltas)
    if deltas:
        cum_drift = deltas[-1] - deltas[0]
        n = len(deltas) - 1 if len(deltas) > 1 else 1
        per_para = cum_drift / n
        max_abs = max(abs(d) for d in deltas)
    else:
        cum_drift = per_para = max_abs = 0.0

    # Detect compensation jumps: deltas going from positive → less positive (or zero)
    # Compute delta-of-delta to find the largest "correction"
    correction_events = []
    for i in range(1, len(deltas)):
        dd = deltas[i] - deltas[i-1]
        if abs(dd) > 5.0:
            correction_events.append({
                'idx': i,
                'word_i': matched[i]['word_i'],
                'word_page': matched[i]['word_page'],
                'oxi_page': matched[i]['oxi_page'],
                'word_y': matched[i]['word_y'],
                'oxi_y': matched[i]['oxi_y'],
                'delta_change': round(dd, 2),
            })

    return {
        'doc_id': data.get('doc_id', '?'),
        'n_matched': len(matched),
        'cum_drift_pt': round(cum_drift, 2),
        'per_para_drift_pt': round(per_para, 4),
        'max_abs_drift_pt': round(max_abs, 2),
        'trajectory_class': trajectory,
        'correction_events': correction_events,
        'first_delta': round(deltas[0], 2) if deltas else 0,
        'last_delta': round(deltas[-1], 2) if deltas else 0,
        'mean_iou': data.get('mean_iou', 0),
        'median_dy': data.get('median_dy', 0),
        'deltas': [round(d, 2) for d in deltas],
    }


def load_phase1_status() -> dict:
    """Load Phase 1 pass/fail per doc from pagination_diff/_summary.json."""
    path = os.path.join(REPO, 'pipeline_data', 'pagination_diff', '_summary.json')
    if not os.path.exists(path):
        return {}
    with open(path, encoding='utf-8') as f:
        s = json.load(f)
    return {d['doc_id']: d for d in s.get('docs', [])}


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    phase1 = load_phase1_status()
    profiles = []
    for path in sorted(glob.glob(os.path.join(IN_DIR, '*.json'))):
        if path.endswith('_summary.json'):
            continue
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        prof = analyze_doc(data)
        # Annotate Phase 1 status
        p1 = phase1.get(prof['doc_id'], {})
        prof['phase1_pass'] = p1.get('pass')
        prof['phase1_score'] = p1.get('score')
        profiles.append(prof)
        # Save per-doc
        out = os.path.join(OUT_DIR, f'{prof["doc_id"]}.json')
        with open(out, 'w', encoding='utf-8') as f:
            json.dump(prof, f, ensure_ascii=False, indent=2)

    # Sort + classify summary
    by_class = defaultdict(list)
    for p in profiles:
        by_class[p['trajectory_class']].append(p['doc_id'])

    summary = {
        'n_total': len(profiles),
        'by_class': {k: {'count': len(v), 'docs': v} for k, v in by_class.items()},
        'top_max_drift': sorted(profiles, key=lambda p: -p['max_abs_drift_pt'])[:15],
        'top_per_para_drift_pos': sorted(profiles, key=lambda p: -p['per_para_drift_pt'])[:10],
        'top_per_para_drift_neg': sorted(profiles, key=lambda p: p['per_para_drift_pt'])[:10],
        'profiles': profiles,
    }
    with open(os.path.join(OUT_DIR, '_summary.json'), 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    # Console report
    print(f'=== Drift Profile Summary (n={len(profiles)}) ===')
    print()
    print('By trajectory class:')
    for k in sorted(by_class.keys(), key=lambda x: -len(by_class[x])):
        docs = by_class[k]
        # Phase 1 status breakdown
        p1_status = {'PASS': 0, 'FAIL': 0, 'unknown': 0}
        for d in docs:
            prof = next(p for p in profiles if p['doc_id'] == d)
            if prof.get('phase1_pass') is True:
                p1_status['PASS'] += 1
            elif prof.get('phase1_pass') is False:
                p1_status['FAIL'] += 1
            else:
                p1_status['unknown'] += 1
        print(f'  {k:<22} {len(docs):>3} docs (PASS={p1_status["PASS"]}, FAIL={p1_status["FAIL"]}, ?={p1_status["unknown"]})')
    print()
    print('Top-15 max-abs drift docs:')
    print(f'  {"doc_id":<32} {"n":>4} {"per_para":>9} {"cum":>8} {"max":>8} {"class":<22} {"P1"}')
    for p in summary['top_max_drift']:
        p1 = 'PASS' if p.get('phase1_pass') else ('FAIL' if p.get('phase1_pass') is False else '?')
        print(f'  {p["doc_id"]:<32} {p["n_matched"]:>4} {p["per_para_drift_pt"]:>+9.4f} {p["cum_drift_pt"]:>+8.2f} {p["max_abs_drift_pt"]:>+8.2f} {p["trajectory_class"]:<22} {p1}')

    print()
    print('Compensating docs (PASS but max_abs > 30pt — load-bearing on bug behavior):')
    comp = [p for p in profiles if p['trajectory_class'] == 'compensating' and p['max_abs_drift_pt'] > 30 and p.get('phase1_pass')]
    for p in comp:
        print(f'  {p["doc_id"]:<32} max={p["max_abs_drift_pt"]:>+7.2f} cum={p["cum_drift_pt"]:>+7.2f} corrections={len(p.get("correction_events", []))}')


if __name__ == '__main__':
    main()
