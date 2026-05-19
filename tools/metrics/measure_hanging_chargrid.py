"""Word COM measurement of hanging-indent + charGrid repros.

For each variant, dump per-char x/y from Word and identify line breaks.
Output JSON for comparison with Oxi dump-layout.
"""
import json
import os
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPRO_DIR = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "tools/metrics/hanging_chargrid_repro")
OUT = os.path.abspath(sys.argv[2] if len(sys.argv) > 2 else "tools/metrics/hanging_chargrid_word.json")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

results = {}
for fname in sorted(os.listdir(REPRO_DIR)):
    if not fname.endswith(".docx"):
        continue
    path = os.path.join(REPRO_DIR, fname)
    label = fname[:-5]
    print(f"\n=== {label} ===")
    doc = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.3)
    para_results = []
    for p_idx in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs(p_idx)
        rng = para.Range
        text = rng.Text.replace("\r", "")
        chars = rng.Characters
        n = min(chars.Count, 200)
        # Per-char measurement: only when y changes do we record line transitions
        last_y = None
        lines = []
        per_char = []
        for c in range(1, n + 1):
            ch = chars(c)
            try:
                x = ch.Information(5)  # wdHorizontalPositionRelativeToPage
                y = ch.Information(6)  # wdVerticalPositionRelativeToPage
            except Exception:
                break
            t = ch.Text
            per_char.append({"c": c, "x": round(x, 3), "y": round(y, 3), "ch": t})
            if last_y is None or abs(y - last_y) > 3:
                lines.append({"char_idx": c, "x": round(x, 3), "y": round(y, 3), "ch": t})
                last_y = y
        # Also extract last char of each line: the char whose y is current but next char's y > current
        line_ends = []
        for i in range(len(per_char) - 1):
            if abs(per_char[i + 1]["y"] - per_char[i]["y"]) > 3:
                line_ends.append(per_char[i])
        if per_char:
            line_ends.append(per_char[-1])
        para_results.append({
            "para_idx": p_idx,
            "text_preview": text[:40],
            "n_chars": chars.Count,
            "lines_start": lines,
            "lines_end": line_ends,
        })
        print(f"  para {p_idx}: n_chars={chars.Count}, lines={len(lines)}")
        for ln, le in zip(lines, line_ends):
            preview = ""
            for pc in per_char:
                if pc["y"] == ln["y"]:
                    preview += pc["ch"]
            print(f"    line@y={ln['y']:.2f}: start x={ln['x']:.2f} {ln['ch']!r}  end x={le['x']:.2f} {le['ch']!r}  N={len(preview)}  preview={preview[:30]!r}...{preview[-12:]!r}")
    results[label] = para_results
    doc.Close(SaveChanges=False)

word.Quit()

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print(f"\nSaved to {OUT}")
