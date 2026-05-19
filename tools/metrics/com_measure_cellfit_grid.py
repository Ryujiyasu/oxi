"""S126 cell-fit COM measurement.

Open each variant in Word, find the table's first cell (restart row of the
vMerge group), and measure per-line char count and pitch.

Method:
  - Use the cell's first paragraph's Range
  - For each line of the rendered paragraph:
    - chars-on-line = count of characters in that line
    - line_y = Information(6) for first char of line
    - pitch = (line_end_x - line_start_x) / (n_chars - 1) approximately

Output: tools/metrics/cellfit_grid/results.json
"""
from __future__ import annotations

import json
import os
import sys
import time
from itertools import product

VARIANTS_DIR = os.path.join(os.path.dirname(__file__), "cellfit_grid", "variants")
OUT_JSON = os.path.join(os.path.dirname(__file__), "cellfit_grid", "results.json")

# Targeted subset for initial diagnostic
TCWS = [1271, 1500, 1800, 2000]
NS = [3, 4, 5, 6, 7, 8]
JCS = ['both', 'left']
BALANCES = [True]


def measure(word, variant_path: str) -> dict:
    doc = word.Documents.Open(os.path.abspath(variant_path), False, True)
    try:
        table = doc.Tables(1)
        cell = table.Cell(1, 1)  # row 1 = vMerge restart
        para = cell.Range.Paragraphs(1).Range
        # Iterate chars; record (char, x, y) via Information(3, 6, 4)
        text = ""
        positions = []  # list of (ch, x, y)
        rng_chars = para.Characters
        n_chars = rng_chars.Count
        for i in range(1, n_chars + 1):
            c = rng_chars(i)
            try:
                x = float(c.Information(3))   # wdHorizontalPositionRelativeToPage
                y = float(c.Information(6))   # wdVerticalPositionRelativeToPage
                ch = c.Text
                positions.append((ch, x, y))
                text += ch
            except Exception:
                pass

        # Group by y to get lines
        lines = {}
        for ch, x, y in positions:
            key = round(y, 1)
            lines.setdefault(key, []).append((x, ch))

        line_info = []
        for y in sorted(lines.keys()):
            items = sorted(lines[y], key=lambda t: t[0])
            xs = [it[0] for it in items]
            chs = ''.join(it[1] for it in items)
            n = len(items)
            if n >= 2:
                pitch = (xs[-1] - xs[0]) / (n - 1)
            else:
                pitch = 0.0
            line_info.append({"y": y, "n": n, "text": chs, "x_first": xs[0] if xs else 0.0,
                               "x_last": xs[-1] if xs else 0.0, "pitch": pitch})

        # Cell width
        col_w_pt = float(cell.Width)
        return {
            "path": os.path.basename(variant_path),
            "text": text,
            "cell_w_pt": col_w_pt,
            "n_lines": len(line_info),
            "lines": line_info,
        }
    finally:
        doc.Close(False)


def main():
    try:
        import win32com.client as com
    except ImportError:
        print("pywin32 required: pip install pywin32")
        sys.exit(1)

    word = com.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone

    results = []
    try:
        for tcw, n, jc, balance in product(TCWS, NS, JCS, BALANCES):
            name = f"cf_tcw{tcw}_n{n}_{jc}_b{int(balance)}.docx"
            path = os.path.join(VARIANTS_DIR, name)
            if not os.path.isfile(path):
                continue
            print(f"  measuring {name}...")
            try:
                t0 = time.time()
                r = measure(word, path)
                r["tcw"] = tcw; r["n_target"] = n; r["jc"] = jc; r["balance"] = balance
                r["elapsed_s"] = round(time.time() - t0, 2)
                results.append(r)
            except Exception as e:
                print(f"    ERR: {e}")
    finally:
        word.Quit()

    with open(OUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"saved {len(results)} measurements to {OUT_JSON}")


if __name__ == "__main__":
    main()
