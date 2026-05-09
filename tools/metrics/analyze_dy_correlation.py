"""Day 32 part 11/12 — analyze per-paragraph correlation CSVs.

Group matched paragraphs by attribute combination and compute mean dy.
Look for which attribute values correlate with non-zero dy.
"""
from __future__ import annotations
import os, sys, csv
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DATA_DIR = os.path.join(REPO, 'pipeline_data')


def load_all():
    rows = []
    for f in os.listdir(DATA_DIR):
        if f.startswith('per_para_correlation_') and f.endswith('.csv'):
            with open(os.path.join(DATA_DIR, f), encoding='utf-8') as fp:
                for row in csv.DictReader(fp):
                    if row['matched'] == 'True':
                        try:
                            row['dy_abs'] = float(row['dy_abs'])
                            row['fs'] = float(row['fs'])
                            row['lh_val'] = float(row['lh_val']) if row['lh_val'] else 0
                            row['lh_rule'] = int(row['lh_rule'])
                            row['snap'] = int(row['snap'])
                            row['text_align'] = int(row['text_align'])
                            row['in_table'] = row['in_table'] == 'True'
                            row['is_empty'] = row['is_empty'] == 'True'
                            rows.append(row)
                        except (ValueError, KeyError):
                            pass
    return rows


def stats(rows, key_fn, label):
    """Group rows by key_fn(row), report mean dy."""
    groups = defaultdict(list)
    for r in rows:
        try:
            k = key_fn(r)
        except Exception:
            continue
        groups[k].append(r['dy_abs'])
    print(f'\n=== {label} ===')
    print(f'  {"key":<40} {"n":>5} {"mean":>8} {"min":>8} {"max":>8}')
    sorted_groups = sorted(groups.items(), key=lambda kv: -abs(sum(kv[1]) / len(kv[1])))
    for k, dys in sorted_groups[:15]:
        mean = sum(dys) / len(dys)
        print(f'  {str(k)[:40]:<40} {len(dys):>5} {mean:>+8.2f} {min(dys):>+8.2f} {max(dys):>+8.2f}')


def main():
    rows = load_all()
    # Filter to same-page only (eliminates cross-page noise)
    rows = [r for r in rows if r.get('oxi_pg') and r.get('word_pg')
            and str(r['oxi_pg']) == str(r['word_pg'])]
    # Recompute dy as page-relative (already is, but verify)
    for r in rows:
        try:
            r['dy_rel'] = float(r['oxi_y']) - float(r['word_y'])
        except (ValueError, KeyError):
            r['dy_rel'] = 0
    # Use dy_rel for analysis
    for r in rows:
        r['dy_abs'] = round(r['dy_rel'], 2)
    print(f'Total same-page matched paragraphs: {len(rows)}')
    docs = set(r['doc_id'] for r in rows)
    print(f'Docs: {sorted(docs)}')

    # By doc
    stats(rows, lambda r: r['doc_id'], 'By doc')

    # By style_name
    stats(rows, lambda r: r['style_name'], 'By style_name')

    # By (doc_id, lh_rule, lh_val)
    stats(rows, lambda r: (r['doc_id'], r['lh_rule'], r['lh_val']), 'By (doc, lh_rule, lh_val)')

    # By (doc_id, fs, in_table)
    stats(rows, lambda r: (r['doc_id'], r['fs'], r['in_table']), 'By (doc, fs, in_table)')

    # By (doc_id, in_table)
    stats(rows, lambda r: (r['doc_id'], r['in_table']), 'By (doc, in_table)')

    # Distribution of dy values per doc
    print('\n=== Per-doc dy distribution ===')
    by_doc = defaultdict(list)
    for r in rows:
        by_doc[r['doc_id']].append(r['dy_abs'])
    for doc_id, dys in sorted(by_doc.items()):
        dys = sorted(dys)
        n = len(dys)
        print(f'  {doc_id}: n={n}, mean={sum(dys)/n:+.2f}, min={dys[0]:+.2f}, '
              f'p25={dys[n//4]:+.2f}, p50={dys[n//2]:+.2f}, p75={dys[3*n//4]:+.2f}, max={dys[-1]:+.2f}')


if __name__ == '__main__':
    main()
