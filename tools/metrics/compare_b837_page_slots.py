"""Compare Word DML vs Oxi layout per-page line slot counts for b837."""
import json, os
from collections import defaultdict

OXI_PATH = os.environ.get("TMP", r"C:\Users\ryuji\AppData\Local\Temp") + r"\b837_layout_fresh.txt"
DML_PATH = r"C:\Users\ryuji\oxi-1\pipeline_data\word_dml\b837808d0555_20240705_resources_data_guideline_02.json"

# --- Parse Oxi layout: TEXT records with (y, page). Cluster y with 5pt tol.
oxi_ys_by_page = defaultdict(list)  # page (1-based) → [y]
current_page = 0
with open(OXI_PATH, encoding='utf-8') as f:
    for line in f:
        line = line.rstrip("\n")
        if line.startswith("PAGE"):
            parts = line.split("\t")
            current_page = int(parts[1]) + 1
        elif line.startswith("TEXT"):
            parts = line.split("\t")
            if len(parts) >= 3:
                try:
                    y = float(parts[2])
                    oxi_ys_by_page[current_page].append(y)
                except ValueError:
                    pass

def cluster(ys, tol=5.0):
    """Cluster close Y values (within `tol`pt) into line positions."""
    if not ys:
        return []
    sorted_ys = sorted(set(round(y, 1) for y in ys))
    clusters = [[sorted_ys[0]]]
    for y in sorted_ys[1:]:
        if y - clusters[-1][-1] <= tol:
            clusters[-1].append(y)
        else:
            clusters.append([y])
    return [sum(c)/len(c) for c in clusters]

oxi_lines_by_page = {p: cluster(ys, tol=10.0) for p, ys in oxi_ys_by_page.items()}

# --- Parse Word DML: walk paragraphs → lines → y ---
with open(DML_PATH, encoding='utf-8') as f:
    dml = json.load(f)

word_lines_by_page = defaultdict(list)
for para in dml.get("paragraphs", []):
    page = para["page"]
    for line in para.get("lines", []):
        word_lines_by_page[page].append(round(line["y"], 1))
# Also include table cells if any
for table in dml.get("tables", []):
    for row in table.get("rows", []):
        for cell in row.get("cells", []):
            for para in cell.get("paragraphs", []):
                page = para.get("page", 0)
                for line in para.get("lines", []):
                    word_lines_by_page[page].append(round(line["y"], 1))

# --- Compare ---
print(f"{'Page':>4}  {'Word':>4}  {'Oxi':>4}  {'Δ':>4}  {'Word top':>9}  {'Oxi top':>9}  {'Word bot':>9}  {'Oxi bot':>9}")
print("-" * 70)
all_pages = sorted(set(oxi_lines_by_page) | set(word_lines_by_page))
for p in all_pages:
    w = sorted(word_lines_by_page[p])
    o = sorted(oxi_lines_by_page[p])
    wt = w[0] if w else -1
    wb = w[-1] if w else -1
    ot = o[0] if o else -1
    ob = o[-1] if o else -1
    print(f"{p:>4}  {len(w):>4}  {len(o):>4}  {len(o)-len(w):>+4}  {wt:>9.1f}  {ot:>9.1f}  {wb:>9.1f}  {ob:>9.1f}")

print("\n--- First line Y per page (Word vs Oxi, first 3 lines) ---")
for p in all_pages:
    w = sorted(word_lines_by_page[p])[:3]
    o = sorted(oxi_lines_by_page[p])[:3]
    wdeltas = [w[i]-w[0] for i in range(len(w))]
    odeltas = [o[i]-o[0] for i in range(len(o))]
    print(f"  p{p}: Word={w} (Δ{wdeltas})  Oxi={[round(x,1) for x in o]} (Δ{[round(x,1) for x in odeltas]})")

print("\n--- Last line Y per page ---")
for p in all_pages:
    w = sorted(word_lines_by_page[p])
    o = sorted(oxi_lines_by_page[p])
    print(f"  p{p}: Word last={w[-1] if w else None}  Oxi last={round(o[-1],1) if o else None}  diff={round((o[-1]-w[-1]) if w and o else 0, 1)}")

print("\n--- Line count delta ---")
cum = 0
for p in all_pages:
    delta = len(oxi_lines_by_page[p]) - len(word_lines_by_page[p])
    cum += delta
    print(f"  p{p}: delta={delta:+d}  cumulative={cum:+d}")

# Map Oxi line Y positions to source content (first 15 chars of the cluster).
print("\n--- Oxi p.4-p.5 line Y positions + first few glyphs ---")
def oxi_lines_with_content(page):
    """Read Oxi layout again to pair y-clusters with chars."""
    records = []
    cur = 0
    pending_text = None
    with open(OXI_PATH, encoding='utf-8') as f:
        for line in f:
            line = line.rstrip("\n")
            if line.startswith("PAGE"):
                cur = int(line.split("\t")[1]) + 1
            elif cur != page:
                continue
            elif line.startswith("TEXT"):
                parts = line.split("\t")
                if len(parts) >= 3:
                    try:
                        y = float(parts[2])
                        x = float(parts[1].split("\t")[0]) if "\t" not in parts[1] else float(parts[1])
                        records.append((y, x, None))
                    except ValueError:
                        pass
            elif line.startswith("T\t") and records:
                # Attach character to last text record
                ch = line[2:]
                y, x, _ = records[-1]
                records[-1] = (y, x, ch)
    # Group by rounded y
    by_y = defaultdict(list)
    for y, x, ch in records:
        if ch:
            by_y[round(y)].append((x, ch))
    rows = []
    for y in sorted(by_y):
        glyphs = sorted(by_y[y], key=lambda g: g[0])
        text = "".join(g[1] for g in glyphs)[:40]
        rows.append((y, text))
    return rows

for p in [4, 5]:
    print(f"  -- Oxi p{p} --")
    for y, t in oxi_lines_with_content(p)[:25]:
        print(f"    y={y}  {t}")

print("\n--- Word p.4-p.5 line Y + text heads ---")
for target_p in [4, 5]:
    print(f"  -- Word p{target_p} --")
    for para in dml.get("paragraphs", []):
        if para["page"] != target_p:
            continue
        for line in para.get("lines", [])[:3]:
            text = para.get("text", "")[:40]
            print(f"    y={line['y']}  {text}")
