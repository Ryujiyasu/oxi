"""drift_profile pre-flight gate for layout fix attempts.

Per Day 31 part 21 strategic conclusion: Bundle fix v9+ requires
pre-flight check to identify which docs would shift > 5pt before
running expensive pipeline.verify (which takes ~10-15 minutes).

Workflow:
1. **Snapshot baseline**: capture current pagination_oxi/<doc>.json for
   each doc as "before" state
2. **Apply fix**: user applies layout code change + rebuilds
3. **Snapshot post-fix**: re-run pagination_oxi → "after" state
4. **Compare**: for each doc, compute MAX-paragraph-y-shift across
   matched paragraphs
5. **Classify**: preserve-class docs (matching+mixed+plateau-pos+comp PASS,
   28 docs) MUST shift ≤ 5pt
6. **Verdict**: PASS if all preserve docs within tolerance; otherwise
   list which docs violate the rule

Output: pipeline_data/drift_preflight/<timestamp>.json with
{doc_id, max_shift_pt, class, ship_class, violation}.

Run from repo root:
  python tools/metrics/drift_profile_preflight.py snapshot baseline
  # ... apply fix, rebuild, run pagination_oxi ...
  python tools/metrics/drift_profile_preflight.py snapshot postfix
  python tools/metrics/drift_profile_preflight.py compare baseline postfix
"""
from __future__ import annotations
import json
import os
import sys
import datetime

# Force UTF-8 output to handle ≤ etc symbols on Windows
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
PAG_OXI_DIR = os.path.join(REPO, 'pipeline_data', 'pagination_oxi')
DRIFT_PROF_DIR = os.path.join(REPO, 'pipeline_data', 'drift_profile')
OUT_DIR = os.path.join(REPO, 'pipeline_data', 'drift_preflight')

# Preserve classes per Day 31 part 8 design (drift_profile classification)
PRESERVE_CLASSES = {'matching', 'mixed', 'plateau-positive', 'plateau-negative'}
# Tolerance: 5pt per CLAUDE.md merge gate proposal
PRESERVE_TOL_PT = 5.0


def load_pagination_oxi(doc_id: str) -> dict | None:
    """Load Oxi pagination data for a doc; returns dict {doc_id, paras: [{para_idx, page, y, x, text}]}."""
    path = os.path.join(PAG_OXI_DIR, f'{doc_id}.json')
    if not os.path.exists(path):
        return None
    with open(path, encoding='utf-8') as f:
        d = json.load(f)
    paras = []
    for page_str, page_paras in d.get('pages', {}).items():
        page = int(page_str)
        for p in page_paras:
            paras.append({
                'para_idx': p.get('para_idx'),
                'page': page,
                'y': p.get('y', 0),
                'x': p.get('x', 0),
                'text': (p.get('text', '') or '')[:30],
            })
    return {'doc_id': doc_id, 'paras': paras}


def snapshot_all() -> dict:
    """Capture current pagination_oxi state for all docs."""
    out = {}
    for fname in sorted(os.listdir(PAG_OXI_DIR)):
        if not fname.endswith('.json') or fname == '_summary.json':
            continue
        doc_id = fname[:-5]
        data = load_pagination_oxi(doc_id)
        if data:
            out[doc_id] = data
    return out


def load_classification() -> dict:
    """Load drift_profile classification: doc_id → trajectory_class."""
    summary_path = os.path.join(DRIFT_PROF_DIR, '_summary.json')
    if not os.path.exists(summary_path):
        return {}
    with open(summary_path, encoding='utf-8') as f:
        s = json.load(f)
    return {p['doc_id']: p.get('trajectory_class', 'unknown') for p in s.get('profiles', [])}


def compute_shift(before: dict, after: dict) -> dict:
    """Compute max paragraph y-shift between before and after snapshots."""
    PAGE_HEIGHT = 841.95
    before_paras = {p['para_idx']: p for p in before['paras'] if p['para_idx'] is not None}
    after_paras = {p['para_idx']: p for p in after['paras'] if p['para_idx'] is not None}
    common_idx = set(before_paras.keys()) & set(after_paras.keys())
    if not common_idx:
        return {'doc_id': before['doc_id'], 'max_shift_pt': 0, 'matched': 0, 'page_count_changed': False}
    max_shift = 0.0
    page_count_changed = False
    for pi in sorted(common_idx):
        b = before_paras[pi]
        a = after_paras[pi]
        if b['page'] != a['page']:
            page_count_changed = True
            # Page change is a large shift
            shift = abs((a['page'] - 1) * PAGE_HEIGHT + a['y'] - ((b['page'] - 1) * PAGE_HEIGHT + b['y']))
        else:
            shift = abs(a['y'] - b['y'])
        if shift > max_shift:
            max_shift = shift
    return {
        'doc_id': before['doc_id'],
        'max_shift_pt': round(max_shift, 2),
        'matched': len(common_idx),
        'page_count_changed': page_count_changed,
    }


def cmd_snapshot(label: str):
    """Snapshot current pagination_oxi to a labeled file."""
    os.makedirs(OUT_DIR, exist_ok=True)
    snap = snapshot_all()
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    out_path = os.path.join(OUT_DIR, f'snap_{label}.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump({'label': label, 'timestamp': timestamp, 'docs': snap}, f, ensure_ascii=False, indent=2)
    print(f'Snapshot {label}: {len(snap)} docs → {out_path}')


def cmd_compare(before_label: str, after_label: str):
    """Compare two snapshots."""
    before_path = os.path.join(OUT_DIR, f'snap_{before_label}.json')
    after_path = os.path.join(OUT_DIR, f'snap_{after_label}.json')
    if not (os.path.exists(before_path) and os.path.exists(after_path)):
        print(f'ERROR: snapshot files not found')
        return 1
    with open(before_path, encoding='utf-8') as f:
        before = json.load(f)
    with open(after_path, encoding='utf-8') as f:
        after = json.load(f)
    classification = load_classification()
    print(f'Comparing {before_label} → {after_label}...')
    print()
    shifts = []
    for doc_id, b_data in before['docs'].items():
        a_data = after['docs'].get(doc_id)
        if not a_data:
            continue
        s = compute_shift(b_data, a_data)
        s['class'] = classification.get(doc_id, 'unknown')
        s['preserve'] = s['class'] in PRESERVE_CLASSES
        s['violation'] = s['preserve'] and s['max_shift_pt'] > PRESERVE_TOL_PT
        shifts.append(s)

    # Categorize
    preserve_violations = [s for s in shifts if s['violation']]
    preserve_ok = [s for s in shifts if s['preserve'] and not s['violation']]
    nonpreserve = [s for s in shifts if not s['preserve']]

    print(f'Preserve docs (target ≤ {PRESERVE_TOL_PT}pt shift): {len(preserve_ok)}/{len(preserve_ok) + len(preserve_violations)} OK')
    print()
    if preserve_violations:
        print(f'[NG] {len(preserve_violations)} preserve-class docs violated tolerance:')
        for s in sorted(preserve_violations, key=lambda x: -x['max_shift_pt']):
            print(f'  {s["doc_id"]:<32} class={s["class"]:<20} max_shift={s["max_shift_pt"]:>+7.2f}pt page_change={s["page_count_changed"]}')
        print()
    print(f'Non-preserve docs (no constraint, top 10 shifts):')
    for s in sorted(nonpreserve, key=lambda x: -x['max_shift_pt'])[:10]:
        print(f'  {s["doc_id"]:<32} class={s["class"]:<20} max_shift={s["max_shift_pt"]:>+7.2f}pt')

    # Save report
    out_path = os.path.join(OUT_DIR, f'compare_{before_label}_vs_{after_label}.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump({
            'before': before_label,
            'after': after_label,
            'preserve_violations': preserve_violations,
            'preserve_ok_count': len(preserve_ok),
            'nonpreserve_top10': sorted(nonpreserve, key=lambda x: -x['max_shift_pt'])[:10],
            'verdict': 'PASS' if not preserve_violations else 'FAIL',
        }, f, ensure_ascii=False, indent=2)
    print()
    print(f'Verdict: {"PASS" if not preserve_violations else "FAIL"}')
    print(f'Report → {out_path}')
    return 0 if not preserve_violations else 1


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    cmd = sys.argv[1]
    if cmd == 'snapshot':
        if len(sys.argv) < 3:
            print('Usage: snapshot <label>')
            return 1
        cmd_snapshot(sys.argv[2])
        return 0
    elif cmd == 'compare':
        if len(sys.argv) < 4:
            print('Usage: compare <before_label> <after_label>')
            return 1
        return cmd_compare(sys.argv[2], sys.argv[3])
    else:
        print(f'Unknown command: {cmd}')
        return 1


if __name__ == '__main__':
    sys.exit(main())
