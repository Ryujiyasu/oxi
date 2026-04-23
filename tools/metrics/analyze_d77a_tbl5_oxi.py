"""Extract Oxi's tbl5 cell content on d77a p.6/p.7 and compare line count
per paragraph against Word's measurement (1/1/1/2/3).

Word measurement (`pipeline_data/d77a_tbl5_word_measurements.json`):
  tbl5 is table_index=5, 1 cell, 5 paragraphs:
    p1: 26 chars, 1 line @ y=649.5
    p2: 28 chars, 1 line @ y=668.5
    p3: 35 chars, 1 line @ y=686.5
    p4: 50 chars, 2 lines @ y=704.5, 722.5
    p5: 74 chars, 3 lines @ y=740.5 (p.6), y=73.5, 109.5 (p.7)  [approx]
  total: 6 lines on p.6 + 2 lines on p.7 = 8 lines

Oxi memory claim: 7 lines total (1 fewer than Word).
"""
import json
from collections import defaultdict
from pathlib import Path


OXI = Path("pipeline_data/d77a_oxi_layout.json")
WORD = Path("pipeline_data/d77a_tbl5_word_measurements.json")


def main():
    oxi = json.load(open(OXI, encoding="utf-8"))
    word = json.load(open(WORD, encoding="utf-8"))

    # Word tbl5 cell paragraphs
    tbl5 = next(t for t in word["tables"] if t["index"] == 5)
    cell = tbl5["cells"][0]
    print("=== Word tbl5 cell paragraphs ===")
    total_w_lines = 0
    for p in cell["paras"]:
        preview = p["preview"][:30]
        print(f"  p{p['p_idx']}: {p['chars']:3d} chars, {p['line_count']} lines")
        for ln in p["lines"]:
            print(f"    pg={ln['page']} y={ln['y_pt']:.2f} '{ln['char']}'")
        total_w_lines += p["line_count"]
    print(f"  Word total: {total_w_lines} lines")

    # Find Oxi's tbl5 on p.6 (1-indexed page 6).
    # tbl5 starts at y~649 (near-bottom of p.6) and continues to p.7.
    # Strategy: find borders on p.6 with y1 >= 640 to identify the cell,
    # then extract text elements within the cell bbox.
    for pi, page in enumerate(oxi["pages"]):
        if page["page"] == 6:
            p6 = page
            break
    else:
        raise SystemExit("No p.6 in Oxi output")

    borders = [e for e in p6["elements"] if e["type"] == "border"]
    # Cell borders are horizontal lines. tbl5 should be near y=640+.
    hb = [b for b in borders if b["h"] == 0.0 and b.get("y1", 0) > 630]
    print("\n=== Oxi p.6 candidate tbl5 borders ===")
    for b in hb:
        print(f"  y1={b.get('y1', '?'):.2f} y2={b.get('y2', '?'):.2f} x1={b.get('x1', '?'):.2f} x2={b.get('x2', '?'):.2f}")

    # Cell bbox: find the tallest border region on p.6 ≥ 640.
    # (Simpler: just dump all p.6 text in y-order, filter to text in y>640 region)
    texts_p6 = [e for e in p6["elements"] if e["type"] == "text" and e["y"] > 630]
    print("\n=== Oxi p.6 text in y>630 region ===")
    for t in sorted(texts_p6, key=lambda e: (e["y"], e["x"])):
        txt = t["text"][:40]
        pi = t.get("para_idx", "?")
        print(f"  pi={pi} x={t['x']:6.2f} y={t['y']:6.2f} fs={t['font_size']} '{txt}'")

    # Group by para_idx and count unique y values
    by_para = defaultdict(set)
    for t in texts_p6:
        pi = t.get("para_idx")
        if pi is not None:
            by_para[pi].add(round(t["y"], 1))
    print("\n=== Oxi p.6 paragraphs in y>630 region (line count = unique y) ===")
    for pi in sorted(by_para.keys()):
        print(f"  para_idx={pi}: {len(by_para[pi])} lines (ys={sorted(by_para[pi])})")

    # p.7: continuation text near top
    for pg in oxi["pages"]:
        if pg["page"] == 7:
            p7 = pg
            break
    texts_p7 = [e for e in p7["elements"] if e["type"] == "text" and e["y"] < 200]
    by_para7 = defaultdict(set)
    for t in texts_p7:
        pi = t.get("para_idx")
        if pi is not None:
            by_para7[pi].add(round(t["y"], 1))
    print("\n=== Oxi p.7 paragraphs in y<200 region ===")
    for pi in sorted(by_para7.keys()):
        print(f"  para_idx={pi}: {len(by_para7[pi])} lines (ys={sorted(by_para7[pi])})")
    print("\n=== Oxi p.7 text in y<200 region ===")
    for t in sorted(texts_p7, key=lambda e: (e["y"], e["x"])):
        txt = t["text"][:40]
        pi = t.get("para_idx", "?")
        print(f"  pi={pi} x={t['x']:6.2f} y={t['y']:6.2f} fs={t['font_size']} '{txt}'")


if __name__ == "__main__":
    main()
