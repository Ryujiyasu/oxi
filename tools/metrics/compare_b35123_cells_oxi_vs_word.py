"""Compare b35123 cell-internal Y: Oxi cached layout vs Word COM measurement."""
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

WORD = Path("pipeline_data/b35123_table_cells_measurement.json")
OXI = Path("pipeline_data/_b35_layout.json")


def main():
    word = json.loads(WORD.read_text(encoding="utf-8"))
    oxi = json.loads(OXI.read_text(encoding="utf-8"))

    print(f"=== b35123 table comparison ===\n")

    # First, find Oxi tables by para_idx grouping
    # Cell elements have para_idx = table's block_idx
    # Group elements by (page, para_idx) to find tables
    oxi_tables = {}  # (page, para_idx) -> {min_y, max_y, n_elements, first_text}
    for page in oxi["pages"]:
        for el in page["elements"]:
            if el.get("type") != "text":
                continue
            pi = el.get("para_idx")
            if pi is None:
                continue
            key = (page["page"], pi)
            if key not in oxi_tables:
                oxi_tables[key] = {"min_y": el["y"], "max_y": el["y"], "n_elements": 0, "first_text": el.get("text", ""), "min_x": el["x"]}
            t = oxi_tables[key]
            t["min_y"] = min(t["min_y"], el["y"])
            t["max_y"] = max(t["max_y"], el["y"])
            t["min_x"] = min(t["min_x"], el["x"])
            t["n_elements"] += 1
            if t["n_elements"] <= 1:
                t["first_text"] = el.get("text", "")

    # Identify likely table containers (large groups with many distinct y values)
    print("Oxi (page, para_idx) groups with > 30 elements:")
    for (page, pi), t in sorted(oxi_tables.items()):
        if t["n_elements"] > 30:
            print(f"  page {page} pidx {pi}: y_range={t['min_y']:.1f}-{t['max_y']:.1f} "
                  f"(span={t['max_y']-t['min_y']:.1f}pt), n_elements={t['n_elements']}, "
                  f"first='{t['first_text'][:20]}'")

    # Now compare to Word's measurements
    print("\nWord tables vs Oxi groups:")
    for tbl in word:
        if "error" in tbl:
            continue
        ti = tbl["table_idx"]
        page = tbl["tbl_page"]
        word_y = tbl["tbl_top_y"]
        # Find Oxi group on same page with min_y closest to word_y
        candidates = [(k, v) for k, v in oxi_tables.items()
                      if k[0] == page and v["n_elements"] > 30]
        if not candidates:
            print(f"  Table {ti} page {page}: Word top_y={word_y}, no Oxi candidate")
            continue
        candidates.sort(key=lambda kv: abs(kv[1]["min_y"] - word_y))
        best = candidates[0]
        oxi_min_y = best[1]["min_y"]
        dy = oxi_min_y - word_y
        print(f"  Table {ti} page {page}: Word_top_y={word_y:.1f}, Oxi_pidx={best[0][1]} min_y={oxi_min_y:.1f}, dy={dy:+.2f}")
        # Per-row first cell comparison
        for row in tbl["rows"][:3]:
            ri = row["row_idx"]
            for cell in row["cells"][:1]:  # first col only
                ci = cell["col_idx"]
                wy = cell["first_char_y"]
                wx = cell["cell_x"]
                # Try to find Oxi text at that x range
                oxi_at_y = []
                for page_obj in oxi["pages"]:
                    if page_obj["page"] != page:
                        continue
                    for el in page_obj["elements"]:
                        if el.get("type") != "text":
                            continue
                        # Match by x range and text
                        if abs(el.get("x", 0) - wx) < 5.0 and abs(el.get("y", 0) - wy) < 25.0:
                            oxi_at_y.append(el)
                if oxi_at_y:
                    oxi_at_y.sort(key=lambda e: abs(e["y"] - wy))
                    closest = oxi_at_y[0]
                    print(f"    Row {ri} col {ci}: Word ({wx:.1f},{wy:.1f}) text='{cell['text'][:15]}' "
                          f"-> Oxi ({closest.get('x'):.1f},{closest.get('y'):.1f}) text='{closest.get('text','')[:5]}' "
                          f"dy={closest['y']-wy:+.2f}")
                else:
                    print(f"    Row {ri} col {ci}: Word ({wx:.1f},{wy:.1f}) text='{cell['text'][:15]}' -> no Oxi match")


if __name__ == "__main__":
    main()
