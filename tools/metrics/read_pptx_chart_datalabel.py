# -*- coding: utf-8 -*-
"""Chart data-labels read: measure chart_datalabel.pdf with fitz — per-page
text spans (data labels / axis / category / title) and vector drawings
(bars / axes / fills) with positions + colors, so we can derive the exact
data-label placement/format/font rules Word uses."""
import sys, os, json
sys.stdout.reconfigure(encoding="utf-8")
import fitz

base = r"pipeline_data\pptx_probes\chart_datalabel"
pdf_path = os.path.join(base, "chart_datalabel.pdf")

doc = fitz.open(pdf_path)
out = {}
for pi in range(len(doc)):
    page = doc[pi]
    d = page.get_text("rawdict")
    spans = []
    for block in d["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                txt = "".join(c["c"] for c in span["chars"])
                if not txt.strip():
                    continue
                o = span["origin"]
                spans.append({
                    "text": txt, "x": round(o[0], 2), "y": round(o[1], 2),
                    "size": round(span["size"], 2), "color": span["color"],
                    "font": span["font"],
                })
    # vector drawings: bars (filled rects), axes (thin lines)
    drawings = []
    for p in page.get_drawings():
        r = p["rect"]
        fill = p.get("fill")
        stroke = p.get("color")
        it = p["items"]
        kinds = [itm[0] for itm in it]
        drawings.append({
            "rect": [round(r.x0, 2), round(r.y0, 2), round(r.x1, 2), round(r.y1, 2)],
            "w": round(r.width, 2), "h": round(r.height, 2),
            "fill": [round(c, 3) for c in fill] if fill else None,
            "stroke": [round(c, 3) for c in stroke] if stroke else None,
            "kinds": kinds[:4],
        })
    out[pi] = {"spans": spans, "drawings": drawings}

with open(os.path.join(base, "chart_datalabel_fitz.json"), "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=1)

# concise console dump
for pi in range(len(doc)):
    print(f"\n===== PAGE {pi+1} =====")
    print("-- spans --")
    for s in out[pi]["spans"]:
        print(f"  '{s['text']}' x={s['x']} y={s['y']} size={s['size']} color=#{s['color']:06x} font={s['font']}")
    print("-- drawings (fill>0 = bars) --")
    for d_ in out[pi]["drawings"]:
        if d_["fill"]:
            print(f"  BAR rect=({d_['rect'][0]},{d_['rect'][1]},{d_['rect'][2]},{d_['rect'][3]}) "
                  f"w={d_['w']} h={d_['h']} fill={[round(c,3) for c in d_['fill']]}")
print("\nsaved:", os.path.join(base, "chart_datalabel_fitz.json"))
