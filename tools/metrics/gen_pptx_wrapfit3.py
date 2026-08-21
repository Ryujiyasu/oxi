# -*- coding: utf-8 -*-
"""wrapfit round 3: verify the master-unit model's sharp predictions.

Rounds 1+2 pinned PowerPoint's fit sum to per-glyph advances rounded to
1/8 pt (the PowerPoint-97 master unit, 576/inch), round-to-nearest, summed --
the unique quantum+mode through all six brackets, and it reproduces d09's
"Happy Holi!" break (master sum 546.5 > box 546.4128 where the float sum
546.399 fits). This round walks 0.005pt steps across each predicted
threshold: the flip must land at the master sum itself if the box width is
compared unquantized, or 1/16pt below it if the width is rounded to master
units too. It also answers inclusive-vs-strict at exact equality.

Usage:
    python tools/metrics/gen_pptx_wrapfit3.py
    python tools/metrics/measure_pptx_word.py pipeline_data/pptx_probes/wrapfit/wrapfit3.pptx pipeline_data/pptx_probes/wrapfit/r3
    python tools/metrics/read_pptx_wrapfit.py 3
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).resolve().parent))
from gen_pptx_wrapfit import FONT_FILES, OUT, add_arm_slide  # noqa: E402

from fontTools.ttLib import TTFont  # noqa: E402
from pptx import Presentation  # noqa: E402


def master_sum_pt(font_path: str, text: str, fs: float) -> float:
    f = TTFont(font_path, lazy=True)
    upm = f["head"].unitsPerEm
    cmap = f.getBestCmap()
    hmtx = f["hmtx"]
    import math
    return sum(
        math.floor(hmtx[cmap[ord(c)]][0] / upm * fs * 8 + 0.5) for c in text
    ) / 8.0


# (label, font, fs, text): windows are computed around the predicted sum.
ARMS = [
    ("arial40", "Arial", 40.0, "west warf"),
    ("arial115", "Arial", 115.64, "west warf"),
    ("segsc40", "Segoe Script", 40.0, "meno mint"),
]


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    manifest = []
    for label, font, fs, text in ARMS:
        s = master_sum_pt(FONT_FILES[font], text, fs)
        # cover [s - 0.075 (one half master unit below), s + 0.02]
        deltas = [round(-0.075 + 0.005 * k, 4) for k in range(20)]
        for delta in deltas:
            add_arm_slide(prs, label, font, fs, text, s + delta)
            manifest.append({
                "slide": len(manifest) + 1,
                "arm": label,
                "font": font,
                "fs": fs,
                "text": text,
                "design_sum_pt": round(s, 4),  # NOTE: master sum, not float
                "delta_pt": delta,
                "width_pt": round(s + delta, 4),
            })
        print(f"{label}: master sum {s}pt, {len(deltas)} slides")
    prs.save(OUT / "wrapfit3.pptx")
    (OUT / "wrapfit3_manifest.json").write_text(
        json.dumps(manifest, indent=1), encoding="utf-8")
    print(f"wrote {OUT / 'wrapfit3.pptx'} ({len(manifest)} slides)")


if __name__ == "__main__":
    main()
