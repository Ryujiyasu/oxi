"""Diff Oxi vs Word layout by matching TEXT content (works around para_idx gaps)."""
# stdout fix for CP932 consoles
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
__doc__ = """Diff Oxi vs Word layout by matching TEXT content (works around para_idx gaps).

Groups Oxi text elements by (page, y_line) → reconstructs lines.
Matches each Oxi line to Word DML's lines by text prefix similarity.
Reports first-mismatch per page to localize the drift source.

Usage: python tools/diff_by_text.py <docid>
Example: python tools/diff_by_text.py d77a
"""
import io, json, os, sys
from collections import defaultdict
from pathlib import Path

DOCID_MAP = {
    "d77a": ("d77a58485f16_20240705_resources_data_outline_08",
             "pipeline_data/_d77a_oxi_layout.json"),
    "683f": ("683ffcab86e2_20230331_resources_open_data_contract_addon_00",
             "pipeline_data/_683f_oxi_layout.json"),
    "0e7a": ("0e7af1ae8f21_20230331_resources_open_data_contract_sample_00",
             "pipeline_data/_0e7a_oxi_layout.json"),
}

def load():
    key = sys.argv[1] if len(sys.argv) > 1 else "d77a"
    full, oxi_path = DOCID_MAP[key]
    word_path = rf"C:/Users/ryuji/oxi-main/pipeline_data/word_dml/{full}.json"
    with io.open(oxi_path, encoding="utf-8") as f: oxi = json.load(f)
    with io.open(word_path, encoding="utf-8") as f: w = json.load(f)
    return full, oxi, w


def oxi_lines_per_page(oxi):
    """Group Oxi text elements into (page_idx, y_key) → line_text concatenated."""
    out = defaultdict(lambda: {"texts": [], "x_min": 1e9, "x_sum": 0, "n": 0})
    for pi, page in enumerate(oxi["pages"]):
        for e in page["elements"]:
            if e["type"] != "text" or not e.get("text"): continue
            # Quantize y to nearest 0.5pt for line grouping
            y_key = round(e["y"] * 2) / 2
            key = (pi + 1, y_key)
            out[key]["texts"].append((e["x"], e["text"]))
            if e["x"] < out[key]["x_min"]: out[key]["x_min"] = e["x"]
    # Sort texts by x, concatenate
    result = {}
    for key, v in out.items():
        v["texts"].sort()
        result[key] = "".join(t for _, t in v["texts"])
    return result, {k: v for k, v in out.items()}


def main():
    full, oxi, w = load()
    oxi_lines, oxi_line_meta = oxi_lines_per_page(oxi)

    # Word lines: each paragraph has a `lines` array with y/x/chars, and `text` (full para text)
    # We need to split the text into lines. DML gives us char counts per line.
    word_lines = []  # list of (page, y, x, text_chars, para_index)
    for p in w["paragraphs"]:
        text = p.get("text", "")
        offset = 0
        for li, line in enumerate(p["lines"]):
            n = line.get("chars", 0)
            chunk = text[offset:offset+n] if n > 0 else ""
            offset += n
            word_lines.append({
                "page": p["page"], "y": line["y"], "x": line["x"],
                "text": chunk, "para": p["index"], "line": li,
            })

    # For each page, find the first Oxi line that doesn't match Word's corresponding line
    print(f"{'page':>4} {'word_y':>7} {'oxi_y':>7} {'Δy':>7} {'word_x':>7} {'oxi_x':>7} {'text':<40}")
    print("-" * 90)
    # Match: for each word line, find an Oxi line on same page with matching text prefix
    seen_oxi = set()
    for wl in word_lines:
        if not wl["text"]: continue  # skip empty (para marker only)
        wtext = wl["text"][:10]
        # Find Oxi line on same page with matching prefix
        match = None
        for (opg, oy), otext in oxi_lines.items():
            if opg != wl["page"]: continue
            if (opg, oy) in seen_oxi: continue
            if otext.startswith(wtext):
                match = ((opg, oy), otext)
                break
        if match:
            (opg, oy), otext = match
            seen_oxi.add((opg, oy))
            ox_min = oxi_line_meta[(opg, oy)]["x_min"]
            dy = oy - wl["y"]
            dx = ox_min - wl["x"]
            mark = " " if abs(dy) < 0.5 else "!"
            print(f"{mark} {wl['page']:>3} {wl['y']:>7.2f} {oy:>7.2f} {dy:>+7.2f} {wl['x']:>7.2f} {ox_min:>7.2f} {wtext:<40}")
        else:
            print(f"? {wl['page']:>3} {wl['y']:>7.2f} {'-':>7} {'-':>7} {wl['x']:>7.2f} {'-':>7} MISSING: {wtext}")


if __name__ == "__main__":
    main()
