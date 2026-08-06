# -*- coding: utf-8 -*-
"""Measure the multi-series LINE chart legend geometry from the Word PDF:
swatches (line + diamond marker via get_drawings items) and labels (rawdict
origins). Reports legend_left / legend_y0 / row pitch vs the single-series
law (legend_y0 = sy+shh/2+17.68, legend_left = plot_right+15.65)."""
import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
import fitz

def merge_lines(lines):
    """Group rawdict chars (over all spans of all lines) into visual lines
    by baseline (|dy|<=0.75pt)."""
    chars = []
    for ln in lines:
        for s in ln.get("spans", []):
            for c in s.get("chars", []):
                chars.append(c)
    chars.sort(key=lambda c: (round(c["origin"][1] * 20) / 20, c["origin"][0]))
    lines = []
    cur = None
    for c in chars:
        y = c["origin"][1]
        if cur is None or abs(y - cur["y"]) > 0.75:
            if cur is not None:
                lines.append(cur)
            cur = {"y": y, "chars": [], "x0": 1e9, "x1": -1e9, "text": ""}
        cur["chars"].append(c)
        cur["x0"] = min(cur["x0"], c["origin"][0])
        cur["x1"] = max(cur["x1"], c["origin"][0] + c.get("adv", 0))
    if cur is not None:
        lines.append(cur)
    for L in lines:
        L["text"] = "".join(c["c"] for c in L["chars"])
    return lines

for name in ("chart_line2", "chart_line3"):
    pdf = os.path.abspath(os.path.join(r"pipeline_data\pptx_probes", name, name + ".pdf"))
    doc = fitz.open(pdf)
    page = doc[0]
    print("=" * 70)
    print(name, "pdf", os.path.getsize(pdf), "page rect", page.rect)
    # --- text labels via rawdict ---
    raw = page.get_text("rawdict")
    lines = merge_lines([ln for b in raw["blocks"] for ln in b.get("lines", [])])
    for L in lines:
        t = L["text"]
        if any(k in t for k in ("Series", "East", "West", "Midwest", "North", "South", "0", "5", "10", "15", "20", "25")):
            print(f"  text x0={L['x0']:.2f} x1={L['x1']:.2f} y={L['y']:.2f}  '{t}'")
    # --- swatches / lines / markers via get_drawings ---
    draws = page.get_drawings()
    for di, d in enumerate(draws):
        rect = d["rect"]
        items = d.get("items", [])
        n_l = sum(1 for it in items if it[0] == "l")
        n_c = sum(1 for it in items if it[0] == "c")
        n_re = sum(1 for it in items if it[0] == "re")
        fill = d.get("fill")
        color = d.get("color")
        # only print drawings in the legend zone (right band) or thin lines
        w = rect.width
        h = rect.height
        if (rect.x0 > 320 and rect.x0 < 520) or (w > 5 and h < 3) or (rect.width < 25 and rect.height < 25):
            print(f"  d{di} rect=({rect.x0:.2f},{rect.y0:.2f},{rect.x1:.2f},{rect.y1:.2f}) "
                  f"w={w:.2f} h={h:.2f} items l={n_l} c={n_c} re={n_re} fill={fill} color={color}")
            for it in items:
                if it[0] == "l":
                    p1, p2 = it[1], it[2]
                    print(f"       line ({p1.x:.2f},{p1.y:.2f})->({p2.x:.2f},{p2.y:.2f})")
                elif it[0] == "c":
                    pts = it[1]
                    print(f"       bezier 4pts: " + " ".join(f"({p.x:.2f},{p.y:.2f})" for p in pts))
                elif it[0] == "re":
                    print(f"       rect {it[1]}")
    doc.close()
