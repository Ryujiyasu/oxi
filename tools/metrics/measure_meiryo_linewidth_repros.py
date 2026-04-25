"""Measure line widths in each Meiryo line-width repro.

For each LW_*.docx, open in Word COM, measure:
  - Number of physical lines in the single paragraph
  - First line: leftmost X, rightmost X, width, char count
  - Per-char dX values to detect compression

Saves results to pipeline_data/meiryo_linewidth_repro.json
"""
import win32com.client
import json
import os
import glob

REPRO_DIR = os.path.abspath("tools/metrics/meiryo_linewidth_repro")
OUT_PATH = "pipeline_data/meiryo_linewidth_repro.json"


def measure_one(word_app, docx_path: str):
    doc = word_app.Documents.Open(docx_path, ReadOnly=True)
    try:
        para = doc.Paragraphs(1)
        r = para.Range
        text = r.Text
        n = len(text)
        per_char = []
        for i in range(n):
            sub = word_app.ActiveDocument.Range(r.Start + i, r.Start + i + 1)
            try:
                x = sub.Information(5)
                y = sub.Information(6)
                per_char.append({'i': i, 'ch': text[i], 'cp': ord(text[i]), 'x': x, 'y': y})
            except:
                per_char.append({'i': i, 'ch': text[i], 'cp': ord(text[i]), 'x': None, 'y': None})

        # Split into physical lines by Y
        lines = []
        cur_line_y = None
        cur_line = []
        for rec in per_char:
            if rec['y'] is None:
                continue
            if cur_line_y is None or abs(rec['y'] - cur_line_y) < 3:
                if cur_line_y is None:
                    cur_line_y = rec['y']
                cur_line.append(rec)
            else:
                lines.append((cur_line_y, cur_line))
                cur_line_y = rec['y']
                cur_line = [rec]
        if cur_line:
            lines.append((cur_line_y, cur_line))

        line_summaries = []
        for (y, chars) in lines:
            # Exclude trailing \r from width calc
            meaningful = [c for c in chars if c['ch'] != '\r']
            if not meaningful:
                continue
            first_x = meaningful[0]['x']
            # Width = last char's advance — use next-char X or line-end
            # For last non-\r char, approximate end via subsequent char's X
            idx_last = meaningful[-1]['i']
            # Look for subsequent char's X (could be \r)
            end_x = None
            for c in chars:
                if c['i'] == idx_last + 1 and c['x'] is not None:
                    end_x = c['x']
                    break
            if end_x is None:
                end_x = meaningful[-1]['x']  # fallback
            line_summaries.append({
                'y': y,
                'first_x': first_x,
                'end_x': end_x,
                'width': end_x - first_x,
                'n_chars': len(meaningful),
                'text': ''.join(c['ch'] for c in meaningful),
            })

        return {
            'text': text.rstrip('\r\x07'),
            'n_chars': len(text.rstrip('\r\x07')),
            'n_lines': len(line_summaries),
            'lines': line_summaries,
            'per_char': per_char,
        }
    finally:
        doc.Close(False)


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    results = {}
    try:
        files = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
        for f in files:
            label = os.path.splitext(os.path.basename(f))[0]
            print(f"\n=== {label} ===")
            try:
                r = measure_one(word, f)
                results[label] = r
                print(f"  text_len={r['n_chars']} n_lines={r['n_lines']}")
                for i, ln in enumerate(r['lines']):
                    print(f"  Line {i+1}: y={ln['y']:.2f} x=[{ln['first_x']:.2f}→{ln['end_x']:.2f}] width={ln['width']:.2f}pt chars={ln['n_chars']}")
            except Exception as e:
                print(f"  ERROR: {e}")
                results[label] = {'error': str(e)}
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {OUT_PATH}")


if __name__ == "__main__":
    main()
