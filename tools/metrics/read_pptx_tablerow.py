# -*- coding: utf-8 -*-
"""Read the table row-height probe back from PowerPoint's PDF.

Table rules are vector rectangles in the PDF, so row heights come out exact to
0.01pt -- far better than detecting them from pixels, which carries about +-1pt
and once led to a false "96dpi quantisation" reading.

Usage: python tools/metrics/read_pptx_tablerow.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_probes" / "tablerow"


def rules(page, min_width: float = 200.0) -> list[float]:
    ys = {
        round((d["rect"].y0 + d["rect"].y1) / 2, 2)
        for d in page.get_drawings()
        if d.get("rect") is not None
        and d["rect"].width > min_width
        and d["rect"].height <= 4
    }
    return sorted(ys)


def main() -> None:
    man = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    pdf = pymupdf.open(ROOT / "probe_tablerow.pdf")
    print(
        f"{'arm':<20}{'sz':>6}{'marT':>8}{'lnSpc':>8}{'face':>9}"
        f"{'row0':>8}{'ctrl':>8}{'net':>8}{'net/sz':>9}{'model':>8}{'err':>7}"
    )
    for m in man:
        ys = rules(pdf[m["slide"] - 1])
        if len(ys) < 4:
            print(f"{m['name']:<20}  rules={ys}")
            continue
        h0 = ys[1] - ys[0]
        ctrl = ys[2] - ys[1]
        net = h0 - 2 * m["marT_pt"]
        pct = (m["lnSpc_pct"] or 100000) / 100000.0
        lines = len((m["text"] or "").split("\n")) if m["text"] else 1
        mult = max(1.2, 1.0625 * pct)
        model = 2 * m["marT_pt"] + mult * m["sz_pt"] * lines
        print(
            f"{m['name']:<20}{m['sz_pt']:>6}{m['marT_pt']:>8.3f}"
            f"{str(m['lnSpc_pct']):>8}{m['typeface']:>9}"
            f"{h0:>8.2f}{ctrl:>8.2f}{net:>8.2f}{net/m['sz_pt']:>9.4f}"
            f"{model:>8.2f}{h0-model:>+7.2f}"
        )


if __name__ == "__main__":
    main()
