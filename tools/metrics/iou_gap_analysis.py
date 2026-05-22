"""S185 Meta: per-doc IoU gap analysis to identify where the
remaining mean IoU debt (0.9400 → 0.9900 = 0.05) actually lives.

For each doc:
  - Compute (1.0 - iou) × n_matched  = "IoU debt" (total paragraph-pts away from perfect)
  - Cross-ref with taxonomy features + positional med_dy
  - Tag bug class by positional dy signature

Tag heuristics:
  TBL_ROW_DRIFT   trH+border with negative med_dy (per-row under-advance)
  CELL_PER_ROW    cell-paras with consistent +2pt drift (tokumei family)
  BODY_OFFSET     consistent +0.5pt med_dy across body paras
  PG_BREAK_CASCADE pos_conf < 0.5 (large page off-by-N cascade)
  LOW_CONF        pos_conf < 0.9 (matcher unreliable, drift unknown)
  TINY            iou >= 0.99 (essentially done)
  MIXED           multiple signatures present

Output: pipeline_data/iou_gap_analysis.json + stdout summary

Usage: python tools/metrics/iou_gap_analysis.py
"""
from __future__ import annotations
import os, sys, json
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

REPO = Path(__file__).resolve().parent.parent.parent
IOU_SUMMARY = REPO / 'pipeline_data' / 'element_iou_diff' / '_summary.json'
POS_SUMMARY = REPO / 'pipeline_data' / 'pagination_diff_positional' / '_summary.json'
TAXONOMY = REPO / 'pipeline_data' / 'doc_feature_taxonomy.json'
OUT_JSON = REPO / 'pipeline_data' / 'iou_gap_analysis.json'


def load_all():
    iou = {}
    if IOU_SUMMARY.exists():
        with open(IOU_SUMMARY, encoding='utf-8') as f:
            d = json.load(f)
        for x in d['docs']:
            iou[x['doc_id']] = x
    pos = {}
    if POS_SUMMARY.exists():
        with open(POS_SUMMARY, encoding='utf-8') as f:
            d = json.load(f)
        for x in d['docs']:
            pos[x['doc_id']] = x
    tax = {}
    if TAXONOMY.exists():
        with open(TAXONOMY, encoding='utf-8') as f:
            d = json.load(f)
        for x in d:
            did = x['doc_id']
            if did not in tax or tax[did].get('iou') is None:
                tax[did] = x
    return iou, pos, tax


def classify(iou_rec, pos_rec, tax_rec) -> str:
    iou_val = iou_rec.get('mean_iou', 1.0)
    if iou_val >= 0.99:
        return 'TINY'
    if not pos_rec:
        return 'NO_POS_DATA'
    conf = pos_rec.get('alignment_confidence', 0)
    med = pos_rec.get('y_diff_visual_median', 0) or 0

    if conf < 0.5:
        return 'PG_BREAK_CASCADE'
    if conf < 0.9:
        return 'LOW_CONF'

    # Classify by med_dy + features
    has_trH = tax_rec.get('n_trH_rows', 0) > 0 if tax_rec else False
    n_tables = tax_rec.get('n_tables', 0) if tax_rec else 0

    if abs(med) < 0.5:
        # high-IoU-debt with small dy: per-paragraph noise or other
        return 'SMALL_DY_OTHER'

    if med < -0.5:
        if has_trH:
            return 'TBL_ROW_DRIFT_NEG'  # 7ead52, a47e family
        return 'BODY_NEG_DRIFT'

    if med > 0.5:
        if n_tables > 0 and has_trH:
            return 'CELL_PER_ROW_POS'  # b35, 29dc6e, b837(tbl) family
        return 'BODY_POS_DRIFT'

    return 'MIXED'


def main():
    iou_by, pos_by, tax_by = load_all()

    rows = []
    for did, iou_rec in iou_by.items():
        if did in ('test', 'pixel', 'gen', 'gen2', 'repro', 'sweep', 'prog'): continue
        iou_val = iou_rec.get('mean_iou', 1.0)
        n_matched = iou_rec.get('n_matched', 0)
        if n_matched == 0:
            continue
        # IoU debt = (1.0 - iou) × n_matched
        debt = (1.0 - iou_val) * n_matched
        pos = pos_by.get(did)
        tax = tax_by.get(did)
        class_tag = classify(iou_rec, pos, tax)
        med_dy = pos.get('y_diff_visual_median') if pos else None
        rows.append({
            'doc_id': did,
            'iou': iou_val,
            'n_matched': n_matched,
            'gap_per_para': 1.0 - iou_val,
            'iou_debt': debt,
            'class': class_tag,
            'pos_conf': pos.get('alignment_confidence') if pos else None,
            'med_dy': med_dy,
            'trH_rows': tax.get('n_trH_rows') if tax else None,
            'n_tables': tax.get('n_tables') if tax else None,
            'grid': tax.get('grid_type') if tax else None,
        })

    # Sort by iou_debt descending
    rows.sort(key=lambda r: -r['iou_debt'])

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    print(f'Wrote {len(rows)} rows to {OUT_JSON}\n')

    # Summary table sorted by debt
    total_debt = sum(r['iou_debt'] for r in rows)
    print(f'Total IoU debt across {len(rows)} real docs: {total_debt:.2f} paragraph-pts')
    print(f'(if all docs at IoU 1.0, debt = 0; current state = {total_debt:.2f})\n')

    print(f'{"doc_id":<14} {"iou":>6} {"n":>5} {"debt":>7} {"med_dy":>7} {"trH":>4} {"grid":<14} {"class":<22}')
    for r in rows[:30]:
        med_s = f'{r["med_dy"]:+.2f}' if r["med_dy"] is not None else '-'
        trh_s = str(r['trH_rows']) if r['trH_rows'] is not None else '-'
        print(f'  {r["doc_id"]:<14} {r["iou"]:.4f} {r["n_matched"]:>5} {r["iou_debt"]:>7.2f} {med_s:>7} {trh_s:>4} {r["grid"] or "-":<14} {r["class"]:<22}')

    print(f'\n=== Debt accounting by class ===')
    from collections import defaultdict
    by_class = defaultdict(lambda: {'docs': 0, 'debt': 0.0})
    for r in rows:
        by_class[r['class']]['docs'] += 1
        by_class[r['class']]['debt'] += r['iou_debt']
    for cls, agg in sorted(by_class.items(), key=lambda kv: -kv[1]['debt']):
        share = 100 * agg['debt'] / total_debt if total_debt > 0 else 0
        print(f'  {cls:<22} docs={agg["docs"]:>3} debt={agg["debt"]:>7.2f}  ({share:5.1f}%)')

    # Top 10 docs by debt
    print(f'\n=== Top 10 docs (highest IoU debt = highest ROI for fix) ===')
    for r in rows[:10]:
        share = 100 * r['iou_debt'] / total_debt if total_debt > 0 else 0
        print(f'  {r["doc_id"]}: iou={r["iou"]:.4f} n={r["n_matched"]:>4} debt={r["iou_debt"]:>6.2f} ({share:5.1f}%) class={r["class"]}')


if __name__ == '__main__':
    main()
