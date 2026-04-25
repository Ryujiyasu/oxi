"""Measure MC_* (multi-font adjacency) repros.

For each doc pattern (観 A B 測) × 10, measure width of A and B across
10 cycles. Summarize per-font behavior.
"""
import win32com.client
import json
import os
import glob

REPRO_DIR = os.path.abspath("tools/metrics/mincho_adjacency_repro")
OUT_PATH = "pipeline_data/mincho_adjacency_widths.json"


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

        # Pattern: 4-char cycle. Extract width of chars at idx 1 and 2 (A, B)
        widths_A = []
        widths_B = []
        n = len(per_char)
        for cycle in range(10):
            base = cycle * 4
            if base + 3 >= n:
                break
            a = per_char[base + 1]
            b = per_char[base + 2]
            c = per_char[base + 3]
            if all(rec['x'] is not None and rec['y'] is not None for rec in [a, b, c]):
                if abs(a['y'] - b['y']) < 3:
                    widths_A.append(b['x'] - a['x'])
                if abs(b['y'] - c['y']) < 3:
                    widths_B.append(c['x'] - b['x'])
        return {'text': text, 'A_widths': widths_A, 'B_widths': widths_B, 'per_char': per_char}
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
            print(f"\n=== {label} ===")
            r = measure_one(word, f)
            results[label] = r
            wa = r['A_widths']
            wb = r['B_widths']
            if wa:
                avg_a = sum(wa)/len(wa)
                print(f"  A char ({len(wa)} samples): avg={avg_a:.3f}pt, values={wa[:5]}")
            if wb:
                avg_b = sum(wb)/len(wb)
                print(f"  B char ({len(wb)} samples): avg={avg_b:.3f}pt, values={wb[:5]}")
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {OUT_PATH}")


if __name__ == "__main__":
    main()
