"""Show all paragraph y diffs (no threshold) for the 5 small-dy 49-doc residuals."""
import json
import subprocess
import os
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OXI_ROOT = os.path.join(os.path.dirname(__file__), "..", "..")
CACHE_DIR = os.path.join(OXI_ROOT, "pipeline_data", "word_dml")

DOCS = [
    "nested_bullet_08",
    "page_break_paragraph_spacing",
    "style_inheritance_complex_19",
    "image_text_wrap_complex_01",
    "mixed_font_line_height",
]

def oxi_paras(docx_path):
    """Run Oxi --structure and return [(para_idx, y, [line_y...])]"""
    result = subprocess.run(
        ["cargo", "run", "--release", "--quiet", "--example", "layout_json", "--", docx_path, "--structure"],
        capture_output=True, text=True, errors="replace", cwd=OXI_ROOT, timeout=120,
    )
    paras = []
    cur = None
    for ln in result.stdout.split("\n"):
        if ln.startswith("PARA\t"):
            if cur: paras.append(cur)
            idx = int(ln.split("\t")[1])
            y = float(ln.split("\t")[2].split("=")[1])
            cur = (idx, y, [])
        elif ln.startswith("  LINE\t"):
            ly = float(ln.split("\t")[1].split("=")[1])
            if cur: cur[2].append(ly)
    if cur: paras.append(cur)
    return paras

for d in DOCS:
    cache = os.path.join(CACHE_DIR, f"{d}.json")
    if not os.path.exists(cache):
        print(f"\n--- {d}: NO CACHE")
        continue
    with open(cache, encoding="utf-8") as f:
        word = json.load(f)
    word_paras = word.get("paragraphs", [])
    docx_path = os.path.join(OXI_ROOT, "pipeline_data", "docx", f"{d}.docx")
    oxi = oxi_paras(docx_path)
    print(f"\n=== {d} ===")
    print(f"Word: {len(word_paras)} paras, Oxi: {len(oxi)} paras")
    for i in range(max(len(word_paras), len(oxi))):
        wy = word_paras[i]["y"] if i < len(word_paras) else None
        oy = oxi[i][1] if i < len(oxi) else None
        if wy is None or oy is None:
            print(f"  P{i}: W={wy} O={oy} (mismatch)")
            continue
        dy = round(oy - wy, 3)
        marker = "  ← " if abs(dy) > 0 else ""
        print(f"  P{i}: W={wy:7.2f} O={oy:7.2f} dy={dy:+6.3f}{marker}")
