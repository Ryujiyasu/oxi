# -*- coding: utf-8 -*-
"""Export the empty-paragraph probe with PowerPoint and read the AAA->BBB gap.

The gap between the two text paragraphs is one body line plus whatever height
PowerPoint gave the empty paragraph between them. Subtracting the body advance
leaves the empty paragraph's own advance, which is what the probe is asking
about.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf
import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\emptypara\emptypara.pptx").resolve()
DST = SRC.with_suffix(".pdf")
BODY_PT = 10.0
ARMS = [
    ("A_epr700", 7.0, 100),
    ("B_epr1000", 10.0, 100),
    ("C_epr2400", 24.0, 100),
    ("D_epr4000", 40.0, 100),
    ("E_none", None, 100),
    ("F_run1000_epr4000", 40.0, 100),
    ("G_epr1000_ls140", 10.0, 140),
    ("H_prev24_none", None, 100),
    ("I_next24_none", None, 100),
    ("J_first_none", None, 100),
]
# The advance that has to be subtracted is the FIRST paragraph's line, which is
# not always the 10pt body: H starts at 24pt, and J has no first paragraph at
# all (its empty line is the box's first, so nothing is subtracted).
FIRST_PT = {"H_prev24_none": 24.0, "J_first_none": 0.0}


def export() -> None:
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        prs = app.Presentations.Open(str(SRC), WithWindow=False)
        try:
            prs.SaveAs(str(DST), 32)
        finally:
            prs.Close()
    finally:
        app.Quit()


# Arms A-I share a first baseline; J has no first paragraph, so its own text
# area top is read from that shared value.
TOP_BASELINE = 84.98


def main() -> None:
    if "--noexport" not in sys.argv:
        export()
    doc = pymupdf.open(DST)
    print(f"{'arm':22s} {'AAA y':>8s} {'BBB y':>8s} {'gap':>8s} {'empty adv':>10s} "
          f"{'/1.2/ls':>9s}  implied pt")
    for i, (label, declared, ls) in enumerate(ARMS):
        d = doc[i].get_text("rawdict")
        ys = {}
        for blk in d["blocks"]:
            for ln in blk.get("lines", []):
                for sp in ln["spans"]:
                    t = "".join(c["c"] for c in sp["chars"]).strip()
                    for key in ("AAA", "BBB"):
                        if t.startswith(key) and key not in ys:
                            ys[key] = sp["chars"][0]["origin"][1]
        if label == "J_first_none":
            ys.setdefault("AAA", TOP_BASELINE)
        if "AAA" not in ys or "BBB" not in ys:
            print(f"{label:22s}  MISSING {sorted(ys)}")
            continue
        gap = ys["BBB"] - ys["AAA"]
        body_adv = FIRST_PT.get(label, BODY_PT) * 1.2 * (ls / 100.0)
        empty_adv = gap - body_adv
        implied = empty_adv / 1.2 / (ls / 100.0)
        dec = "default" if declared is None else f"{declared:g}pt"
        print(f"{label:22s} {ys['AAA']:8.2f} {ys['BBB']:8.2f} {gap:8.2f} {empty_adv:10.2f} "
              f"{implied:9.2f}  declared {dec}")


if __name__ == "__main__":
    main()
