"""Measure bracket-pair repro widths via Word COM.

For each BP_*.docx, compute per-char dX to get actual rendered widths.
"""
import win32com.client
import json
import os
import glob

REPRO_DIR = os.path.abspath("tools/metrics/bracket_pair_repro")
OUT_PATH = "pipeline_data/bracket_pair_widths.json"


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

        # Group by line
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

        # Per-char width (dX) within each line
        widths_by_line = []
        for (y, chars) in lines:
            cws = []
            for j in range(len(chars) - 1):
                w = chars[j+1]['x'] - chars[j]['x']
                cws.append({'i': chars[j]['i'], 'ch': chars[j]['ch'],
                            'cp': chars[j]['cp'], 'x': chars[j]['x'], 'width': w})
            if chars:
                cws.append({'i': chars[-1]['i'], 'ch': chars[-1]['ch'],
                            'cp': chars[-1]['cp'], 'x': chars[-1]['x'], 'width': None})
            widths_by_line.append({'y': y, 'widths': cws})
        return {
            'text': text.rstrip('\r'),
            'n_lines': len(lines),
            'lines': widths_by_line,
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
                # Print per-char widths for line 1
                if r['lines']:
                    line1 = r['lines'][0]
                    print(f"  Line 1 y={line1['y']:.2f}, n_chars={len(line1['widths'])}")
                    # Show first 30 chars with widths
                    for rec in line1['widths'][:40]:
                        ch = rec['ch']
                        cp = rec['cp']
                        w = rec['width']
                        w_s = f"{w:.2f}" if w is not None else "--"
                        ch_s = 'LF' if ch == '\n' else ('CR' if ch == '\r' else ch)
                        print(f"    [{rec['i']:3d}] '{ch_s}' U+{cp:04X} x={rec['x']:.2f} w={w_s}")
                    # Summary: count of compressed chars (<10pt)
                    ws = [r['width'] for r in line1['widths'] if r['width'] is not None]
                    if ws:
                        total = sum(ws)
                        print(f"  Line 1 total ink: {total:.2f}pt, avg: {total/len(ws):.3f}pt/char ({len(ws)} chars)")
                        n_compressed = sum(1 for w in ws if w < 9.5)
                        print(f"  Compressed (< 9.5pt): {n_compressed} / {len(ws)}")
                if len(r['lines']) > 1:
                    line2 = r['lines'][1]
                    print(f"  Line 2 y={line2['y']:.2f}, n_chars={len(line2['widths'])}")
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
