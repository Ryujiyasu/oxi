"""Compare Word vs Oxi line heights on d77a p.7 body paragraphs.

If cursor_y fix cascades to page count 13, one hypothesis is that Oxi's line
heights on p.7 body are slightly different from Word's, causing accumulated
drift through p.7..p.10 even when the first para position matches.
"""
import json
from collections import defaultdict
from pathlib import Path


OXI = Path("pipeline_data/d77a_oxi_layout.json")
WORD_ALL = Path("pipeline_data/d77a_all_paras_measurement.json")


def main():
    oxi = json.load(open(OXI, encoding="utf-8"))
    word = json.load(open(WORD_ALL, encoding="utf-8"))

    # Word p.7 body paragraphs
    w_p7 = [p for p in word["paras"] if p["page"] == 7 and not p["in_table"]]
    print("=== Word p.7 body paragraphs ===")
    for i, p in enumerate(w_p7):
        print(f"  idx={p['idx']:3d} y={p['y_pt']:6.2f} {'[E]' if p['is_empty'] else '[B]'} '{p['text'][:40]}'")
    if len(w_p7) >= 2:
        print("  --- gaps ---")
        for i in range(1, len(w_p7)):
            gap = w_p7[i]['y_pt'] - w_p7[i-1]['y_pt']
            print(f"    idx={w_p7[i]['idx']:3d} gap={gap:+.2f}pt")

    # Oxi p.7 content: group by para_idx, take min y per para as paragraph-top
    for pg in oxi["pages"]:
        if pg["page"] == 7:
            p7 = pg
            break

    by_para = defaultdict(list)
    for e in p7["elements"]:
        if e["type"] == "text":
            pi = e.get("para_idx")
            if pi is not None:
                by_para[pi].append(e["y"])
    print("\n=== Oxi p.7 paragraphs (min y per para) ===")
    # Sort by min y
    paras_sorted = sorted(by_para.items(), key=lambda kv: min(kv[1]))
    last_y = None
    for pi, ys in paras_sorted:
        y_min = min(ys)
        unique_ys = sorted(set(round(y, 1) for y in ys))
        line_count = len(unique_ys)
        gap = (y_min - last_y) if last_y else 0
        print(f"  pi={pi:3d} y={y_min:6.2f} lines={line_count} ys={unique_ys[:3]}{'...' if len(unique_ys)>3 else ''}"
              + (f"  gap={gap:+.2f}" if last_y else ""))
        last_y = y_min


if __name__ == "__main__":
    main()
