# -*- coding: utf-8 -*-
"""wrapfit round 2: how big is the epsilon, and does it scale with font size?

Round 1 bracketed PowerPoint's fit threshold inside (design_sum, design_sum
+0.15pt] at 40pt with NO trailing-ink term. This deck walks the first 0.15pt
in 0.01pt steps at three font sizes; if the flip point scales with fs the
epsilon is an em fraction (a fixed-point guard in the layout engine), if it
stays put it is an absolute length. d09's knife-edge (breaks at +0.0138pt
slack, 115.64pt) is the case any candidate must reproduce.

Usage:
    python tools/metrics/gen_pptx_wrapfit2.py
    python tools/metrics/measure_pptx_word.py pipeline_data/pptx_probes/wrapfit/wrapfit2.pptx <tmpdir>  (copy deck.pdf back as deck2.pdf)
    python tools/metrics/read_pptx_wrapfit.py 2
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).resolve().parent))
from gen_pptx_wrapfit import FONT_FILES, OUT, add_arm_slide, design_sum_pt  # noqa: E402

from pptx import Presentation  # noqa: E402

ARMS = [
    ("arial40",  "Arial", 40.0,   "west warf", 0.0, 0.15, 0.01),
    ("arial115", "Arial", 115.64, "west warf", 0.0, 0.45, 0.03),
    ("arial20",  "Arial", 20.0,   "west warf", 0.0, 0.08, 0.005),
]


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    manifest = []
    for label, font, fs, text, lo, hi, step in ARMS:
        s = design_sum_pt(FONT_FILES[font], text, fs)
        n = int(round((hi - lo) / step)) + 1
        for k in range(n):
            delta = lo + k * step
            add_arm_slide(prs, label, font, fs, text, s + delta)
            manifest.append({
                "slide": len(manifest) + 1,
                "arm": label,
                "font": font,
                "fs": fs,
                "text": text,
                "design_sum_pt": round(s, 4),
                "delta_pt": round(delta, 4),
                "width_pt": round(s + delta, 4),
            })
        print(f"{label}: design sum {s:.3f}pt at {fs}pt, {n} slides")
    prs.save(OUT / "wrapfit2.pptx")
    (OUT / "wrapfit2_manifest.json").write_text(
        json.dumps(manifest, indent=1), encoding="utf-8")
    print(f"wrote {OUT / 'wrapfit2.pptx'} ({len(manifest)} slides)")


if __name__ == "__main__":
    main()
