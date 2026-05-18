"""S105 v2 analysis: classify rows by XML hRule + check formula per class."""
import json, math, sys
from pathlib import Path
from collections import Counter

sys.stdout.reconfigure(encoding='utf-8')


def main():
    with open('c:/Users/ryuji/oxi-main/tools/metrics/atleast_snap_survey_v2.json', encoding='utf-8') as f:
        rows = json.load(f)
    print(f"Total rows: {len(rows)}")

    # Classify
    classes = Counter()
    for r in rows:
        rule = r.get('h_rule_xml_effective')
        has = r.get('has_trH_xml')
        if not has:
            classes['no_trH_element'] += 1
        elif rule is None:
            classes['has_trH_no_rule'] += 1  # shouldn't happen
        else:
            classes[f'has_trH_{rule}'] += 1
    print("\n=== Row classification by XML ===")
    for k, v in classes.most_common():
        print(f"  {k}: {v}")

    # Filter to rows WITH explicit trHeight + measurable pitch
    explicit = [r for r in rows if r.get('has_trH_xml') and r.get('rendered_pitch_pt') is not None and r.get('trH_xml_pt')]
    print(f"\nRows with explicit trH + measurable pitch: {len(explicit)}")

    # Sub-classify by hRule
    by_rule = {}
    for r in explicit:
        rule = r.get('h_rule_xml_effective') or 'auto'
        by_rule.setdefault(rule, []).append(r)
    print("\n=== Rule breakdown ===")
    for rule, rs in by_rule.items():
        print(f"  {rule}: {len(rs)}")

    # For each rule, check what's the relationship between trH and pitch
    print("\n=== Per-rule pitch vs trH analysis ===")
    for rule, rs in by_rule.items():
        print(f"\n--- rule = {rule} ({len(rs)} rows) ---")
        # Categorize: pitch = trH? pitch < trH? pitch > trH (snap)?
        equal = 0
        less = 0
        snap_match = 0  # pitch = ceil((trH + 0.5)/0.75) * 0.75
        more_no_match = 0
        samples_diff = []
        for r in rs:
            trh = r['trH_xml_pt']
            pitch = r['rendered_pitch_pt']
            bw = r.get('border_pt', 0.5)
            # Border may come as raw sz (4 = 0.5pt) or pt — try both
            bw_pt = bw / 8 if bw > 2 else bw
            snap_pred = math.ceil((trh + bw_pt) / 0.75) * 0.75
            d = pitch - trh
            d_snap = pitch - snap_pred
            samples_diff.append((trh, pitch, snap_pred, d, d_snap))
            if abs(pitch - trh) < 0.5:
                equal += 1
            elif pitch < trh - 0.5:
                less += 1
            elif abs(d_snap) < 0.5:
                snap_match += 1
            else:
                more_no_match += 1
        print(f"  pitch ≈ trH (matches exact): {equal}")
        print(f"  pitch < trH (Word ignores atLeast): {less}")
        print(f"  pitch ≈ S98 snap formula: {snap_match}")
        print(f"  pitch > trH but no snap match: {more_no_match}")

        # Print samples
        print(f"  Top 10 samples (sorted by |d_snap|):")
        for trh, pitch, snap, d, ds in sorted(samples_diff, key=lambda x: abs(x[4]))[:10]:
            print(f"    trH={trh:6.2f} pitch={pitch:7.2f} S98_pred={snap:7.2f} d_trH={d:+.2f} d_S98={ds:+.2f}")


if __name__ == '__main__':
    main()
