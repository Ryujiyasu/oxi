"""Diff Oxi layout positions vs Word DML positions for 0e7a p.2.

Reads:
- pipeline_data/_0e7a_oxi_layout.json (from --dump-layout)
- C:/Users/ryuji/oxi-main/pipeline_data/word_dml/0e7a...json

Identifies per-paragraph x/y drift on p.2 to localize the remaining bug
(post-line-wrap fix; p.2 is now rank 1 bottom-5 at 0.5767).
"""
import io
import json

OXI = "pipeline_data/_0e7a_oxi_layout.json"
WORD = r"C:/Users/ryuji/oxi-main/pipeline_data/word_dml/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.json"

with io.open(OXI, encoding="utf-8") as f:
    oxi = json.load(f)
with io.open(WORD, encoding="utf-8") as f:
    w = json.load(f)

# Word paragraphs for p.2
word_p2 = [p for p in w["paragraphs"] if p["page"] == 2]
print(f"Word p.2: {len(word_p2)} paragraphs")
print(f"  first 5: {[(p['index'], p['x'], p['y']) for p in word_p2[:5]]}")
print(f"  last 5:  {[(p['index'], p['x'], p['y']) for p in word_p2[-5:]]}")

# Oxi elements on p.2 — aggregate per paragraph (take MIN x and MIN y per para_idx)
oxi_p2 = oxi["pages"][1]["elements"]
oxi_text = [e for e in oxi_p2 if e["type"] == "text" and e["para_idx"] is not None]
by_para = {}
for e in oxi_text:
    pi = e["para_idx"]
    if pi not in by_para or e["y"] < by_para[pi]["y"]:
        by_para[pi] = {"x": e["x"], "y": e["y"]}
    elif e["y"] == by_para[pi]["y"] and e["x"] < by_para[pi]["x"]:
        by_para[pi]["x"] = e["x"]

print(f"\nOxi p.2 unique paragraph_index count: {len(by_para)}")
oxi_pairs = sorted(by_para.items())
print(f"  first 5: {[(pi, round(v['x'],2), round(v['y'],2)) for pi, v in oxi_pairs[:5]]}")
print(f"  last 5:  {[(pi, round(v['x'],2), round(v['y'],2)) for pi, v in oxi_pairs[-5:]]}")

# Word para indexes are 1-based (per DML extractor); Oxi is 0-based.
# Map: Oxi para_idx i → Word para index i+1? Or different offset?
# Compare first Oxi para on p.2 with first Word para on p.2.
word_first_on_p2 = word_p2[0]
oxi_first_on_p2 = oxi_pairs[0]
print(f"\nWord first on p.2: idx={word_first_on_p2['index']} x={word_first_on_p2['x']} y={word_first_on_p2['y']}")
print(f"Oxi  first on p.2: idx={oxi_first_on_p2[0]} x={oxi_first_on_p2[1]['x']:.2f} y={oxi_first_on_p2[1]['y']:.2f}")

# Try offset: Word idx ≈ Oxi idx + K
# Scan first 30 word p.2 paras, find best K
print("\nLooking for index offset...")
w_indices = [p["index"] for p in word_p2[:30]]
o_indices = [pi for pi, _ in oxi_pairs[:30]]
print(f"  Word indices: {w_indices[:10]}")
print(f"  Oxi  indices: {o_indices[:10]}")
if w_indices and o_indices:
    k = w_indices[0] - o_indices[0]
    print(f"  inferred offset K = {k} (Word_idx = Oxi_idx + {k})")

# Per-para y diff on p.2 (up to 30 paragraphs)
print(f"\n{'i':>4} {'word_idx':>9} {'oxi_idx':>8} {'word_y':>7} {'oxi_y':>7} {'Δy':>7} {'word_x':>7} {'oxi_x':>7} {'Δx':>7}")
print("-" * 84)
for i in range(min(30, len(oxi_pairs), len(word_p2))):
    wi = word_p2[i]
    oi, ov = oxi_pairs[i]
    dy = ov["y"] - wi["y"]
    dx = ov["x"] - wi["x"]
    print(f"{i:>4} {wi['index']:>9} {oi:>8} {wi['y']:>7.2f} {ov['y']:>7.2f} {dy:>+7.2f} {wi['x']:>7.2f} {ov['x']:>7.2f} {dx:>+7.2f}")
