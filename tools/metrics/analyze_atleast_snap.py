"""S104 analysis: derive refined atLeast snap formula from COM survey data.

Hypothesis space:
A. S98 formula: pitch = ceil((max(trH, content_h) + bw) / 0.75) × 0.75
B. trH-only: pitch = ceil((trH + bw) / 0.75) × 0.75 when trH > content
C. No snap: pitch = max(trH, content_h)
D. Other context-dependent

For each surveyed row, compute predicted_A and predicted_C, then compare
to actual rendered pitch. Classify each row by which prediction matches.
"""
import json
import sys
from pathlib import Path
from collections import Counter

sys.stdout.reconfigure(encoding='utf-8')

SURVEY = Path('c:/Users/ryuji/oxi-main/tools/metrics/atleast_snap_survey.json')


def snap_formula(row_h_pt: float, border_pt: float) -> float:
    return (((row_h_pt + border_pt) / 0.75).__ceil__() * 0.75)


def main():
    with open(SURVEY, encoding='utf-8') as f:
        rows = json.load(f)

    print(f"Surveyed {len(rows)} atLeast rows across baseline")
    # Skip rows without measurable pitch
    rows_with_pitch = [r for r in rows if r.get('rendered_pitch_pt') is not None]
    print(f"  {len(rows_with_pitch)} with measurable pitch")

    # Classify each row by which formula fits
    classes = Counter()
    by_doc = {}
    detail = []
    for r in rows_with_pitch:
        trH = r.get('trH_pt') or 0
        bw = r.get('border_pt') or 0.5
        actual = r['rendered_pitch_pt']

        # Predictions
        pred_A = snap_formula(trH, bw) if trH > 0 else None  # S98 formula applied to trH alone
        # Need content_h estimate — approximate via font_size if available
        fs = r.get('font_size_pt') or 10.5
        # Single-line natural height ≈ 1.4 × fs for CJK
        content_h_est = 1.4 * fs

        pred_max = max(trH, content_h_est)
        pred_B = snap_formula(pred_max, bw)  # S98 formula with max
        pred_C = pred_max  # no snap

        # Which is closer?
        def diff(p): return abs(actual - p) if p is not None else 999
        scores = {
            'A_trH_only_snap': diff(pred_A),
            'B_max_snap': diff(pred_B),
            'C_no_snap': diff(pred_C),
        }
        best = min(scores, key=scores.get)
        classes[best] += 1
        by_doc.setdefault(r['doc'], Counter())[best] += 1
        detail.append({
            **r,
            'pred_A_trH_snap': pred_A,
            'pred_B_max_snap': pred_B,
            'pred_C_no_snap': pred_C,
            'best': best,
            'best_diff': scores[best],
        })

    print("\n=== Classification across all rows ===")
    for k, v in classes.most_common():
        print(f"  {k}: {v} ({v/len(rows_with_pitch)*100:.1f}%)")

    print("\n=== Top 10 docs by row count ===")
    for doc, c in sorted(by_doc.items(), key=lambda x: -sum(x[1].values()))[:10]:
        total = sum(c.values())
        breakdown = ', '.join(f'{k.split("_")[0]}={v}' for k, v in c.most_common())
        print(f"  {doc[:50]}: total={total} | {breakdown}")

    # Show distribution of pred_B (S98 formula) diff
    pred_B_diffs = [d['best_diff'] if d['best'] == 'B_max_snap' else abs(d['rendered_pitch_pt'] - d['pred_B_max_snap']) for d in detail]
    near_match_B = sum(1 for x in pred_B_diffs if x < 0.5)
    moderate_B = sum(1 for x in pred_B_diffs if 0.5 <= x < 2.0)
    far_B = sum(1 for x in pred_B_diffs if x >= 2.0)
    print(f"\nPred B (S98 formula) accuracy:")
    print(f"  <0.5pt diff (match): {near_match_B}")
    print(f"  0.5-2.0pt: {moderate_B}")
    print(f"  >=2.0pt (poor): {far_B}")

    # Look at rows where formula B is WRONG: what's different?
    bad_B = [d for d in detail if d['best'] != 'B_max_snap' and abs(d['rendered_pitch_pt'] - (d['pred_B_max_snap'] or 0)) > 1.0]
    print(f"\n=== Rows where formula B is wrong (>1pt off, top 15 examples) ===")
    print(f"{'doc':<40} {'trH':>6} {'rendered':>10} {'pred_B':>8} {'best':<15} {'n_cells':>3} {'n_paras':>3} {'vA':>3}")
    for d in sorted(bad_B, key=lambda x: -abs(x['rendered_pitch_pt'] - (x['pred_B_max_snap'] or 0)))[:15]:
        doc_short = d['doc'][:38]
        print(f"{doc_short:<40} {d['trH_pt'] or 0:>6.1f} {d['rendered_pitch_pt']:>10.2f} {d['pred_B_max_snap'] or 0:>8.2f} {d['best']:<15} {d['n_cells']:>3} {d['n_paras_per_cell']:>3} {d['v_align']!s:>3}")

    out = Path(SURVEY.parent / 'atleast_snap_analysis.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump({
            'classes': dict(classes),
            'by_doc': {k: dict(v) for k, v in by_doc.items()},
            'detail': detail[:200],  # truncate for readability
            'pred_B_accuracy': {'match_lt0.5': near_match_B, '0.5_to_2': moderate_B, 'gt_2': far_B},
        }, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {out}")


if __name__ == '__main__':
    main()
