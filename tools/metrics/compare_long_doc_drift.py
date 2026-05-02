"""Compare Word COM-measured Y vs Oxi cached layout Y for long-doc paragraphs.

Pin: does Oxi accumulate Y drift over many paragraphs?
"""
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

WORD = Path("pipeline_data/long_doc_drift_measurement.json")

# Map doc_id → cached Oxi layout path
LAYOUTS = {
    "e3c545": "pipeline_data/_e3c545_layout.json",
    # 04b88, 34140b have no cached Oxi layout
}


def load_oxi_y_per_para(layout_path: Path) -> dict:
    """Return {(page, para_idx_0based): y} from Oxi cached layout."""
    d = json.loads(layout_path.read_text(encoding="utf-8"))
    out = {}
    for page in d["pages"]:
        page_num = page["page"]
        for el in page["elements"]:
            if el.get("type") != "text":
                continue
            pi = el.get("para_idx")
            if pi is None:
                continue
            y = el.get("y")
            # First text on each para: store min y per para_idx
            key = pi  # use just para_idx; Oxi may not have page-aware idx
            if key not in out or y < out[key][1]:
                out[key] = (page_num, y)
    return out


def main():
    word_data = json.loads(WORD.read_text(encoding="utf-8"))
    print(f"{'doc':<8} {'p_idx':>5} {'Word page':>5} {'Word y':>7}  {'Oxi page':>5} {'Oxi y':>7}  {'dy':>6}")
    for doc in word_data:
        doc_id = doc["doc_id"]
        if "error" in doc or doc_id not in LAYOUTS:
            print(f"  skip {doc_id} (no oxi layout)")
            continue
        oxi_y = load_oxi_y_per_para(Path(LAYOUTS[doc_id]))
        for s in doc["samples"]:
            if "error" in s:
                continue
            # Word p_i is 1-based; Oxi para_idx is 0-based
            pi = s["i"] - 1
            oxi = oxi_y.get(pi)
            if oxi is None:
                print(f"{doc_id:<8} {s['i']:>5} {s['page']:>5} {s['y']:>7.1f}  {'-':>5} {'-':>7}  N/A")
                continue
            page_o, y_o = oxi
            dy = round(y_o - s["y"], 2)
            print(f"{doc_id:<8} {s['i']:>5} {s['page']:>5} {s['y']:>7.1f}  {page_o:>5} {y_o:>7.1f}  {dy:>+6.2f}")


if __name__ == "__main__":
    main()
