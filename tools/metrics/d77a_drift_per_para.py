"""Compare Word COM paragraph y_start vs Oxi's para-start-y for d77a.
Find where cumulative Y drift first appears and grows.
"""
import json
import subprocess
from pathlib import Path

ROOT = Path(__file__).parent.parent.parent
COM = ROOT / "pipeline_data" / "d77a_p6_p7_cascade.json"
OXI = Path(r"c:/tmp/d77a_main.txt")

# Parse Word COM
with open(COM, encoding="utf-8") as f:
    word = json.load(f)

# Parse Oxi layout
oxi_text = OXI.read_text(encoding="utf-8", errors="replace").splitlines()
oxi_pages = {}
cur = -1
for l in oxi_text:
    if l.startswith("PAGE"):
        cur = int(l.split("\t")[1])
        oxi_pages[cur] = []
    elif cur >= 0 and l.startswith("TEXT"):
        parts = l.split("\t")
        oxi_pages[cur].append({
            "x": float(parts[1]), "y": float(parts[2]),
            "w": float(parts[3]), "h": float(parts[4]),
        })

# Oxi page numbers are 0-indexed; Oxi has 13 pages vs Word's 12.
# Report Oxi's first text of each unique y cluster on each page.
# Group by (page, y_rounded) to find line starts.

def first_x_per_line(page_elems):
    """Per unique y, find smallest x."""
    by_y = {}
    for e in page_elems:
        yr = round(e["y"], 1)
        if yr not in by_y or e["x"] < by_y[yr]:
            by_y[yr] = e["x"]
    return sorted(by_y.items())


print(f'{"Word":20} {"":15} {"Oxi":20} {"Δy":>6}')
print("-" * 80)
# Print paragraph-by-paragraph: each Word para has y_start + page_start.
# Oxi mapping: since total line count may match, align by text content.
# Simplest: compare Word para y_start to Oxi's line whose starting X/shape
# matches the corresponding position.
# For quick analysis, just dump Word para y and find Oxi line at similar y.

# Build Oxi y list per page
oxi_ys_per_page = {}
for pi, elems in oxi_pages.items():
    oxi_ys_per_page[pi+1] = sorted({round(e["y"], 1) for e in elems})

# Word has 12 pages. Oxi has 13. Offset unknown. Try both alignments.
# Heuristic: find per-page closest Oxi y to each Word y.

for para in word:
    wp = para["page_start"]
    wy = para["y_start"]
    if not isinstance(wy, (int, float)): continue
    if wy < 0: continue
    text = para["text"][:20]
    # Find closest Oxi line on either Oxi page wp or wp+1 (offset by 1)
    best = None
    for opg_candidate in (wp, wp + 1):
        if opg_candidate not in oxi_ys_per_page: continue
        for oy in oxi_ys_per_page[opg_candidate]:
            dy = oy - wy
            if best is None or abs(dy) < abs(best[2]):
                best = (opg_candidate, oy, dy)
    if best:
        print(f'Wp{wp:2} y={wy:6.1f}  {text!r:30}  Op{best[0]} y={best[1]:6.1f} Δy={best[2]:+.1f}')
