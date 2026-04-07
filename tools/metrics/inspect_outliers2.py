"""Show paragraph metrics around the dy boundaries for the 5 outliers."""
import json, sys
sys.stdout.reconfigure(encoding="utf-8")

DOCS = [
    "nested_bullet_08",
    "page_break_paragraph_spacing",
    "style_inheritance_complex_19",
    "image_text_wrap_complex_01",
    "mixed_font_line_height",
]

for d in DOCS:
    print(f"\n=== {d} ===")
    with open(f"pipeline_data/word_dml/{d}.json", encoding="utf-8") as f:
        data = json.load(f)
    for p in data["paragraphs"]:
        ls = p.get("line_spacing")
        sa = p.get("space_after")
        sb = p.get("space_before")
        font = (p.get("font") or "")[:14]
        sz = p.get("font_size")
        nl = len(p.get("lines", []))
        text = (p.get("text") or "")[:30]
        print(f"  P{p['index']:2d} y={p['y']:7.2f} ls={ls} sa={sa} sb={sb} font={font:14s} sz={sz} nl={nl} text={text!r}")
