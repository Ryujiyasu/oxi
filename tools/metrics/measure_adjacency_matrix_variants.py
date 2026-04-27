"""Measure adjacency_matrix_repro_{V_*} variants and compare to baseline.

Reads the 3 variant repro dirs created by `build_adjacency_matrix_variants.py`,
measures per-char advance widths via Word COM (same logic as
`measure_adjacency_matrix.py`), and writes:
  - pipeline_data/adjacency_matrix_widths_V_NOFE.json
  - pipeline_data/adjacency_matrix_widths_V_CP.json
  - pipeline_data/adjacency_matrix_widths_V_COMPAT15.json

Then prints a 2-col diff table per variant vs baseline
(`adjacency_matrix_widths.json` from 2026-04-25):
  - "matches baseline"  → that knob is NOT the trigger
  - "differs (full)"    → that knob is the (or a) trigger; removing it
                          eliminates compression
  - "differs (compr)"   → removing the knob STILL compresses; trigger is
                          something else still in baseline (rare)

Run after build_adjacency_matrix_variants.py. Requires Word + win32com.
"""
import glob
import json
import os

import win32com.client


# V_NOFE removed 2026-04-27 — useFELayout falsified as gate by LW_30/LW_31 per-char
# compare in meiryo_linewidth_repro.json (identical 5.50pt 、「 compression).
VARIANTS = ['V_CP', 'V_COMPAT15']
BASELINE_JSON = 'pipeline_data/adjacency_matrix_widths.json'

PUNCTS = ['CM', 'PD', 'LBK', 'RBK', 'LPN', 'RPN', 'FPD', 'FCM']
CH = {'CM': '、', 'PD': '。', 'LBK': '「', 'RBK': '」',
      'LPN': '（', 'RPN': '）', 'FPD': '．', 'FCM': '，'}


def measure_one(word_app, docx_path):
    doc = word_app.Documents.Open(docx_path, ReadOnly=True)
    try:
        para = doc.Paragraphs(1)
        r = para.Range
        text = r.Text
        per_char = []
        for i in range(len(text)):
            sub = word_app.ActiveDocument.Range(r.Start + i, r.Start + i + 1)
            try:
                x = sub.Information(5)
                y = sub.Information(6)
                per_char.append({'i': i, 'ch': text[i], 'x': x, 'y': y})
            except Exception:
                per_char.append({'i': i, 'ch': text[i], 'x': None, 'y': None})
        return per_char
    finally:
        doc.Close(False)


def widths_from_per_char(per_char):
    widths = {'prev': [], 'next': []}
    n = len(per_char)
    for cycle in range(10):
        base = cycle * 4
        if base + 2 >= n:
            break
        prev_rec = per_char[base + 1]
        next_rec = per_char[base + 2]
        post_rec = per_char[base + 3] if base + 3 < n else None
        if prev_rec['x'] is None or next_rec['x'] is None:
            continue
        if prev_rec['y'] is None or next_rec['y'] is None:
            continue
        if abs(prev_rec['y'] - next_rec['y']) > 3:
            continue
        widths['prev'].append(next_rec['x'] - prev_rec['x'])
        if post_rec and post_rec['x'] is not None and abs(post_rec['y'] - next_rec['y']) < 3:
            widths['next'].append(post_rec['x'] - next_rec['x'])
    return widths


def measure_dir(word, repro_dir):
    results = {}
    files = sorted(glob.glob(os.path.join(repro_dir, '*.docx')))
    for f in files:
        label = os.path.splitext(os.path.basename(f))[0]
        per_char = measure_one(word, f)
        results[label] = widths_from_per_char(per_char)
    return results


def avg(xs):
    return sum(xs) / len(xs) if xs else None


def classify(variant_avg, baseline_avg, full=10.5):
    if variant_avg is None or baseline_avg is None:
        return '--'
    # tolerance: 10tw grid means values land on n/2 boundaries; use 0.3 slop
    if abs(variant_avg - baseline_avg) < 0.3:
        return 'match'
    if variant_avg > full - 0.3:
        return 'full'   # no compression in variant
    return 'diff'


def print_diff_matrix(label, variant_results, baseline_results, axis):
    print(f'\n== {label} :: {axis} char width vs baseline ==')
    print(f"{'P\\N':>4s}", end='')
    for n in PUNCTS:
        print(f" {CH[n]:>5s}", end='')
    print()
    for p in PUNCTS:
        print(f"{CH[p]:>4s}", end='')
        for n in PUNCTS:
            lbl = f'ADJ_{p}_{n}'
            v = variant_results.get(lbl, {}).get(axis, [])
            b = baseline_results.get(lbl, {}).get(axis, [])
            va = avg(v)
            ba = avg(b)
            tag = classify(va, ba)
            if tag == '--':
                print(f" {'--':>5s}", end='')
            elif tag == 'match':
                print(f" {va:5.2f}", end='')
            elif tag == 'full':
                print(f" {va:5.2f}*", end='')  # * = differs (full = no compression)
            else:
                print(f" {va:5.2f}!", end='')  # ! = differs but still compressed
        print()


def main():
    with open(BASELINE_JSON, 'r', encoding='utf-8') as f:
        baseline = json.load(f)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        for vid in VARIANTS:
            repro_dir = os.path.abspath(f'tools/metrics/adjacency_matrix_repro_{vid}')
            if not os.path.isdir(repro_dir):
                print(f'[{vid}] missing dir {repro_dir} - run build_adjacency_matrix_variants.py first')
                continue
            print(f'[{vid}] measuring {repro_dir} ...')
            results = measure_dir(word, repro_dir)
            out_path = f'pipeline_data/adjacency_matrix_widths_{vid}.json'
            os.makedirs(os.path.dirname(out_path), exist_ok=True)
            with open(out_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
            print(f'[{vid}] saved {out_path}')
            print_diff_matrix(vid, results, baseline, 'prev')
            print_diff_matrix(vid, results, baseline, 'next')
    finally:
        word.Quit()

    print()
    print('Legend: bare value = matches baseline within 0.3pt')
    print('        value*     = differs from baseline AND ~ full width (compression LOST in variant)')
    print('        value!     = differs from baseline AND still compressed')
    print()
    print('Interpretation:')
    print('  V_CP cells matching baseline       -> cSC=cP does not change closing-punct compression rate')
    print('  V_CP * cells                       -> cP gives EXTRA compression (beyond baseline)')
    print('  V_COMPAT15 cells matching baseline -> compat=14/15 same yakumono behaviour')
    print('  V_COMPAT15 * cells                 -> compat=14 had stricter rule, compat=15 relaxes')
    print()
    print('  If both variants match baseline:')
    print('    -> Word applies next-trigger rule UNCONDITIONALLY for compat>=14, cSC in')
    print('       {doNotCompress, compressPunctuation}, useFELayout {on, off}, kern {on, off}.')
    print('       Oxi mod.rs:4140 gate `compress_punctuation` is over-restrictive and')
    print('       blocks Word-correct behaviour for the ~all baseline docs using')
    print('       cSC=doNotCompress. See RESEARCH_LOG.md 2026-04-27.')


if __name__ == "__main__":
    main()
