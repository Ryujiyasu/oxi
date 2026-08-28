# -*- coding: utf-8 -*-
"""Read the rendered pitch of each lnSpc-rounding arm out of the probe PDF."""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).resolve().parent))
from gen_pptx_lnspcround import ARMS  # noqa: E402


def main() -> None:
    pdf = pymupdf.open(Path(r"pipeline_data\pptx_probes\lnspcround\lnspcround.pdf").resolve())
    print(f"{'declared':>9} {'rendered':>9} {'delta':>7}   {'floor':>5} {'round':>5} {'half':>5}")
    for i, val in enumerate(ARMS):
        dec = val / 100
        ys = sorted({round(l["bbox"][1], 2)
                     for b in pdf[i].get_text("dict")["blocks"] if b["type"] == 0
                     for l in b["lines"]})
        ys = [y for y in ys if y > min(ys) + 5] if len(ys) > 6 else ys
        runs, cur = [], [ys[0]]
        for y in ys[1:]:
            (cur if abs(y - cur[-1] - (cur[1] - cur[0] if len(cur) > 1 else y - cur[-1])) < 1.5
             else runs.append(cur) or cur.clear() or cur).append(y) if False else None
            if len(cur) < 2 or abs((y - cur[-1]) - (cur[1] - cur[0])) < 1.0:
                cur.append(y)
            else:
                runs.append(cur)
                cur = [y]
        runs.append(cur)
        run = max(runs, key=len)
        p = float(np.polyfit(range(len(run)), run, 1)[0])
        ok = lambda v: "OK" if abs(v - p) < 0.1 else ""
        print(f"{dec:9.2f} {p:9.3f} {p - dec:+7.3f}   {ok(int(dec)):>5} {ok(round(dec)):>5} "
              f"{ok(round(dec * 2) / 2):>5}   ({len(run)} lines)")


if __name__ == "__main__":
    main()
