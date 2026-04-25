"""Measure compression in all 60 adjacency repros.

For each ADJ_<prev>_<next>.docx, text = (観 prev next 測) × 10.
We expect per-char width 10.5 for all unless compressed.

Report for each repro:
  - prev_width (avg across 10 occurrences)
  - next_width (avg across 10 occurrences)
  - whether compression detected (< 9pt)
"""
import win32com.client
import json
import os
import glob

REPRO_DIR = os.path.abspath("tools/metrics/adjacency_matrix_repro")
OUT_PATH = "pipeline_data/adjacency_matrix_widths.json"


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
            except:
                per_char.append({'i': i, 'ch': text[i], 'x': None, 'y': None})
        return per_char
    finally:
        doc.Close(False)


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    results = {}
    try:
        files = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
        for f in files:
            label = os.path.splitext(os.path.basename(f))[0]
            per_char = measure_one(word, f)
            # Compute widths (X[i+1] - X[i]) for consecutive chars on same line
            # Pattern: 観 prev next 測 (×10) — first char is at idx 0 (line start)
            # We want widths at idx 1 (prev) and idx 2 (next), repeated cycles
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
                    continue  # cycle crosses line boundary
                prev_w = next_rec['x'] - prev_rec['x']
                widths['prev'].append(prev_w)
                if post_rec and post_rec['x'] is not None and abs(post_rec['y'] - next_rec['y']) < 3:
                    next_w = post_rec['x'] - next_rec['x']
                    widths['next'].append(next_w)
            results[label] = widths
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    # Print matrix
    PUNCTS = ['CM', 'PD', 'LBK', 'RBK', 'LPN', 'RPN', 'FPD', 'FCM']
    CH = {'CM':'、', 'PD':'。', 'LBK':'「', 'RBK':'」',
          'LPN':'（', 'RPN':'）', 'FPD':'．', 'FCM':'，'}

    print("\n=== Prev char width (avg) ===")
    print(f"{'prev\\next':>10s}", end='')
    for n in PUNCTS:
        print(f" {CH[n]:>5s}", end='')
    print()
    for p in PUNCTS:
        print(f"{CH[p]:>10s}", end='')
        for n in PUNCTS:
            label = f'ADJ_{p}_{n}'
            if label in results and results[label]['prev']:
                avg = sum(results[label]['prev']) / len(results[label]['prev'])
                print(f" {avg:5.2f}", end='')
            else:
                print(f" {'--':>5s}", end='')
        print()

    print("\n=== Next char width (avg) ===")
    print(f"{'prev\\next':>10s}", end='')
    for n in PUNCTS:
        print(f" {CH[n]:>5s}", end='')
    print()
    for p in PUNCTS:
        print(f"{CH[p]:>10s}", end='')
        for n in PUNCTS:
            label = f'ADJ_{p}_{n}'
            if label in results and results[label]['next']:
                avg = sum(results[label]['next']) / len(results[label]['next'])
                print(f" {avg:5.2f}", end='')
            else:
                print(f" {'--':>5s}", end='')
        print()

    print(f"\nSaved to {OUT_PATH}")


if __name__ == "__main__":
    main()
