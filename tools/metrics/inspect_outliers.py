"""Show all paragraph dy mismatches for the 5 P|dy| outlier docs."""
import json, sys, subprocess, os
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dml_diff import get_oxi_structure, get_word_structure

DOCS = [
    "nested_bullet_08",
    "page_break_paragraph_spacing",
    "style_inheritance_complex_19",
    "image_text_wrap_complex_01",
    "mixed_font_line_height",
]

for d in DOCS:
    print(f"\n=== {d} ===")
    cache = f"pipeline_data/word_dml/{d}.json"
    docx = f"pipeline_data/docx/{d}.docx"
    word = get_word_structure(cache)
    oxi = get_oxi_structure(docx)
    for pi in range(min(len(word["pages"]), len(oxi["pages"]))):
        w_paras = [p for p in word["pages"][pi]["paragraphs"] if p.get("lines")]
        o_paras = oxi["pages"][pi]["paragraphs"]
        used = set()
        for wp in w_paras:
            best_oi = None; best_dy = float("inf")
            for oi, op in enumerate(o_paras):
                if oi in used: continue
                dy = abs(op["y"] - wp["y"])
                if dy < best_dy: best_dy = dy; best_oi = oi
            if best_oi is None: continue
            used.add(best_oi)
            op = o_paras[best_oi]
            dy = op["y"] - wp["y"]
            wls = wp.get("lines", [])
            ols = op.get("lines", [])
            wl1 = wls[0] if wls else None
            ol1 = ols[0] if ols else None
            l1dy = (ol1["y"] - wl1["y"]) if (wl1 and ol1) else 0
            print(f"  P{wp['index']:2d} W.y={wp['y']:7.2f} O.y={op['y']:7.2f} dy={dy:+5.2f}  L1dy={l1dy:+5.2f}  W.lines={len(wls)} O.lines={len(ols)}")
