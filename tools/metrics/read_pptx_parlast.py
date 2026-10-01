# -*- coding: utf-8 -*-
"""Read the parlast probe: every line's baseline, PowerPoint vs Oxi.

Prints, per arm, each line's baseline offset from the box top in PowerPoint's
PDF and the Oxi - PowerPoint difference (from `--dump-layout`), so a
last-line-only error shows as a column that is zero except in its last cell.

    python tools/metrics/read_pptx_parlast.py [--env K=V]
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
DIR = REPO / "pipeline_data" / "pptx_probes" / "parlast"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"
WORDS = ["HAMBURG", "MINDEN", "BAXTER"]


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--env", action="append", default=[])
    args = ap.parse_args()
    meta = json.loads((DIR / "arms.json").read_text(encoding="utf-8"))
    top = meta["box_top_pt"]
    doc = pymupdf.open(DIR / "probe_parlast.pdf")
    env = dict(os.environ)
    env.update(kv.split("=", 1) for kv in args.env)
    with tempfile.TemporaryDirectory() as tmp:
        dump = Path(tmp) / "dump.json"
        subprocess.run([str(EXE), str(DIR / "probe_parlast.pptx"), str(Path(tmp) / "s"), "72",
                        f"--dump-layout={dump}"], capture_output=True, env=env, check=False)
        oxi = json.loads(dump.read_text(encoding="utf-8"))
    worst = 0.0
    for index, arm in enumerate(meta["arms"]):
        spans = {}
        for block in doc[index].get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                for span in line["spans"]:
                    word = span["text"].strip()
                    if word in WORDS and word not in spans:
                        spans[word] = span["origin"][1] - top
        truth = [spans.get(w) for w in WORDS[:arm["lines"]]]
        paras = [p for sh in oxi["slides"][index]["shapes"]
                 for p in (sh.get("content") or {}).get("paragraphs", [])]
        got = [b - top for b in paras[0]["line_baselines"]] if paras else []
        diffs = []
        for k, t in enumerate(truth):
            if t is None or k >= len(got):
                diffs.append("   n/a")
                continue
            diffs.append(f"{got[k] - t:+6.2f}")
            worst = max(worst, abs(got[k] - t))
        ts = "  ".join("   n/a" if t is None else f"{t:6.2f}" for t in truth)
        print(f"{arm['font'][:19]:20s} {arm['pct']:3d}% {arm['mode']:4s} L{arm['lines']} "
              f"oxi_lines={len(got)}  ppt [{ts}]  oxi-ppt [{'  '.join(diffs)}]")
    print(f"\nworst |oxi - ppt| = {worst:.3f}pt")


if __name__ == "__main__":
    main()
