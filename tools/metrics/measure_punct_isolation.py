"""Measure per-punct width in Word COM for each PI_* repro.

For each doc, find per-char X positions and compute:
  - Width of regular CJK chars (baseline)
  - Width of the single punctuation char at each occurrence (10 samples each)
  - Avg & distribution of punct width
  - Whether punct at line-start / line-end behaves differently

Save results and print summary table.
"""
import win32com.client
import json
import os
import glob

REPRO_DIR = os.path.abspath("tools/metrics/punct_isolation_repro")
OUT_PATH = "pipeline_data/punct_isolation_widths.json"


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
                per_char.append({'i': i, 'ch': text[i], 'cp': ord(text[i]), 'x': x, 'y': y})
            except:
                per_char.append({'i': i, 'ch': text[i], 'cp': ord(text[i]), 'x': None, 'y': None})

        # Compute per-line structure
        lines = []
        cur_y = None
        cur = []
        for rec in per_char:
            if rec['y'] is None:
                continue
            if cur_y is None or abs(rec['y'] - cur_y) < 3:
                if cur_y is None:
                    cur_y = rec['y']
                cur.append(rec)
            else:
                lines.append((cur_y, cur))
                cur_y = rec['y']
                cur = [rec]
        if cur:
            lines.append((cur_y, cur))

        # For each line, compute width of each char: width[i] = X[i+1] - X[i]
        # If X[i+1] is on next line, use line-end fallback (skip)
        widths_by_line = []
        for (y, chars) in lines:
            char_widths = []
            for j in range(len(chars) - 1):
                w = chars[j+1]['x'] - chars[j]['x']
                char_widths.append({'i': chars[j]['i'], 'ch': chars[j]['ch'], 'cp': chars[j]['cp'],
                                    'x': chars[j]['x'], 'width': w})
            # last char's width unknown (goes to next line or \r)
            if len(chars) >= 1:
                char_widths.append({'i': chars[-1]['i'], 'ch': chars[-1]['ch'], 'cp': chars[-1]['cp'],
                                    'x': chars[-1]['x'], 'width': None})
            widths_by_line.append({'y': y, 'widths': char_widths})

        # Stats: for each non-\r char, how wide?
        return {
            'text': text,
            'n_lines': len(lines),
            'lines_ink_widths': widths_by_line,
            'per_char': per_char,
        }
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
            try:
                r = measure_one(word, f)
                results[label] = r

                # Compute stats
                total_widths = []
                punct_widths = []  # width of the special punct char
                cjk_widths = []  # width of regular CJK
                for line in r['lines_ink_widths']:
                    for rec in line['widths']:
                        if rec['width'] is None:
                            continue
                        if rec['ch'] in '観測値定':
                            cjk_widths.append(rec['width'])
                        elif rec['ch'] == '\r':
                            pass
                        else:
                            # This is the "special" punct char for this repro
                            punct_widths.append(rec['width'])
                        total_widths.append(rec['width'])

                cjk_avg = sum(cjk_widths)/len(cjk_widths) if cjk_widths else 0
                punct_avg = sum(punct_widths)/len(punct_widths) if punct_widths else 0
                print(f"  CJK chars ({len(cjk_widths)}): avg={cjk_avg:.3f}pt, widths[0:3]={cjk_widths[:3]}")
                if punct_widths:
                    print(f"  PUNCT chars ({len(punct_widths)}): avg={punct_avg:.3f}pt, widths[0:5]={punct_widths[:5]}")
                print(f"  Lines: {r['n_lines']}, total chars: {sum(len(ln['widths'])-1 for ln in r['lines_ink_widths'])}")

            except Exception as e:
                print(f"  ERROR: {e}")
                import traceback; traceback.print_exc()
                results[label] = {'error': str(e)}
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {OUT_PATH}")


if __name__ == "__main__":
    main()
