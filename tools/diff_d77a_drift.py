"""Find where d77a p.9 drift begins by diffing Oxi vs Word DML across all pages."""
import io, json
from collections import defaultdict

OXI = "pipeline_data/_d77a_oxi_layout.json"
WORD = r"C:/Users/ryuji/oxi-main/pipeline_data/word_dml/d77a58485f16_20240705_resources_data_outline_08.json"

with io.open(OXI, encoding="utf-8") as f:
    oxi = json.load(f)
with io.open(WORD, encoding="utf-8") as f:
    w = json.load(f)

# First paragraph per page (Word) and equivalent Oxi paragraph
word_by_page = defaultdict(list)
for p in w["paragraphs"]:
    word_by_page[p["page"]].append(p)

# Oxi: first para_idx per page (min paragraph_index on each page)
oxi_first_per_page = {}
for pi, page in enumerate(oxi["pages"]):
    text_els = [e for e in page["elements"] if e["type"] == "text" and e["para_idx"] is not None]
    if text_els:
        min_pi = min(e["para_idx"] for e in text_els)
        oxi_first_per_page[pi + 1] = min_pi  # page is 1-based

print("Per-page first paragraph (Word idx vs Oxi idx, drift):")
print(f"{'page':>4} {'word_first':>10} {'oxi_first':>9} {'drift_paras':>11} {'word_y':>7} {'oxi_y':>7}")
print("-" * 60)
# Infer Oxi-to-Word offset using page 1
p1_word_min = min(p["index"] for p in word_by_page[1]) if word_by_page[1] else 0
p1_oxi_min = oxi_first_per_page.get(1, 0)
k = p1_word_min - p1_oxi_min
print(f"(offset k = {k}; Word_idx = Oxi_idx + {k})")

for pg in sorted(word_by_page.keys()):
    word_first = min(p["index"] for p in word_by_page[pg])
    word_first_y = min(p["y"] for p in word_by_page[pg] if p["index"] == word_first)
    oxi_first = oxi_first_per_page.get(pg)
    oxi_first_y = None
    if oxi_first is not None:
        for e in oxi["pages"][pg - 1]["elements"]:
            if e.get("para_idx") == oxi_first and e["type"] == "text":
                if oxi_first_y is None or e["y"] < oxi_first_y:
                    oxi_first_y = e["y"]
    # Expected Oxi idx if no drift: word_first - k
    expected_oxi = word_first - k
    drift = (oxi_first or 0) - expected_oxi  # positive = Oxi has more paras per page
    yw = f"{word_first_y:.2f}" if word_first_y is not None else "—"
    yo = f"{oxi_first_y:.2f}" if oxi_first_y is not None else "—"
    print(f"{pg:>4} {word_first:>10} {str(oxi_first):>9} {drift:>+11} {yw:>7} {yo:>7}")

# Where does drift first appear?
print("\nDrift appears when:")
prev_drift = 0
for pg in sorted(word_by_page.keys()):
    word_first = min(p["index"] for p in word_by_page[pg])
    oxi_first = oxi_first_per_page.get(pg, 0)
    drift = oxi_first - (word_first - k)
    if drift != prev_drift:
        print(f"  page {pg}: drift jumped {prev_drift} → {drift}")
        prev_drift = drift
