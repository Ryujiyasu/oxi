# -*- coding: utf-8 -*-
"""Read the sub-100% line-advance probe back out of PowerPoint's PDF.

Prints, per arm, the measured 2nd->3rd baseline advance beside the two rival
models: a flat `1.2 * fs * n`, and the face's own natural line height with 1.2
as a floor. See gen_pptx_lineadv.py for why the faces straddle 1.2.

Usage: python tools/metrics/read_pptx_lineadv.py
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pymupdf
from fontTools.ttLib import TTFont

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_probes" / "lineadv"
FILES = {
    "Arial": "arial.ttf", "Times New Roman": "times.ttf", "Georgia": "georgia.ttf",
    "Verdana": "verdana.ttf", "Comic Sans MS": "comic.ttf", "Segoe Script": "segoesc.ttf",
}


def face_metrics(name: str) -> dict:
    p = Path(os.environ["WINDIR"]) / "Fonts" / FILES[name]
    f = TTFont(p, lazy=True, checkChecksums=0)
    u = f["head"].unitsPerEm
    hh, os2 = f["hhea"], f["OS/2"]
    return {
        "hhea": (hh.ascent - hh.descent) / u,
        "hhea_gap": (hh.ascent - hh.descent + hh.lineGap) / u,
        "typo": (os2.sTypoAscender - os2.sTypoDescender) / u,
        "typo_gap": (os2.sTypoAscender - os2.sTypoDescender + os2.sTypoLineGap) / u,
        "win": (os2.usWinAscent + os2.usWinDescent) / u,
    }


def main() -> None:
    man = json.loads((OUT / "manifest.json").read_text(encoding="utf-8"))
    pdf = pymupdf.open(OUT / "probe_lineadv.pdf")
    met = {f: face_metrics(f) for f in FILES}
    keys = ["hhea", "hhea_gap", "typo", "typo_gap", "win"]
    print(f"{'arm':<24}{'sz':>8}{'n':>6}{'measured':>10}{'flat 1.2':>10}"
          + "".join(f"{k:>11}" for k in keys))
    resid = {k: [] for k in keys + ["flat"]}
    for m in man:
        pg = pdf[m["slide"] - 1]
        ys = []
        for b in pg.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                ch = [c for s in l["spans"] for c in s["chars"]]
                if ch:
                    ys.append(ch[0]["origin"][1])
        ys.sort()
        if len(ys) < 3:
            print(f"{m['name']:<24} only {len(ys)} baselines")
            continue
        adv = ys[2] - ys[1]
        fs, n = m["sz_pt"], m["lnSpc_pct"] / 100000.0
        flat = 1.2 * fs * n
        row = f"{m['name']:<24}{fs:>8.2f}{n:>6.2f}{adv:>10.2f}{flat:>10.2f}"
        resid["flat"].append(adv - flat)
        for k in keys:
            v = max(1.2, met[m["typeface"]][k]) * fs * n
            resid[k].append(adv - v)
            row += f"{v:>11.2f}"
        print(row)
    print(f"\n{'model':<12}{'mean|err|':>12}{'max|err|':>12}   (pt)")
    for k in ["flat"] + keys:
        e = [abs(v) for v in resid[k]]
        if e:
            print(f"{k:<12}{sum(e) / len(e):>12.3f}{max(e):>12.3f}")
    print("\nface ratios (max(1.2, r) is what the model multiplies):")
    for f, d in met.items():
        print(f"  {f:<16}" + "  ".join(f"{k}={d[k]:.4f}" for k in keys))


if __name__ == "__main__":
    main()
