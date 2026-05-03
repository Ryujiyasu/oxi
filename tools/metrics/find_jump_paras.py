"""For each doc, find the EXACT paragraphs where Δy jumps by ~10pt (one wrap-line worth)."""
import json, glob, os, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DIR = "C:/Users/ryuji/AppData/Local/Temp/diffv2_out"
for f in sorted(glob.glob(os.path.join(DIR, "*.json"))):
    with open(f, encoding='utf-8') as fh:
        d = json.load(fh)
    doc_id = d["doc_id"]
    matches = d["matches"]
    print(f"\n=== {doc_id[:50]} jump points (Δy increase >= 5pt over previous matched para) ===")
    prev_y = None
    for m in matches:
        dy = m["divergence"]["y_diff"]
        if dy is None: continue
        if prev_y is not None:
            jump = dy - prev_y
            if jump >= 5:
                wp = m["word"]
                print(f"  P{wp['idx']:>3} '{wp['text'][:50]:50s}' Δy_now={dy:>+6.2f} (was {prev_y:>+6.2f}, jump=+{jump:.2f})")
        prev_y = dy
