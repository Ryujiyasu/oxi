# -*- coding: utf-8 -*-
"""Read the paired empty-paragraph probe: WITH minus WITHOUT is the advance."""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf
import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\emptypara2\emptypara2.pptx").resolve()
DST = SRC.with_suffix(".pdf")
QUESTIONS = [
    ("ctl_10_10_epr1000", 10.0, 10.0, 1000),
    ("prev24_next10_none", 24.0, 10.0, None),
    ("prev10_next24_none", 10.0, 24.0, None),
    ("prev32_next10_none", 32.0, 10.0, None),
    ("prev10_next10_none", 10.0, 10.0, None),
    ("prev24_next10_epr1000", 24.0, 10.0, 1000),
]


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


def gap(page) -> float:
    ys = {}
    for blk in page.get_text("rawdict")["blocks"]:
        for ln in blk.get("lines", []):
            for sp in ln["spans"]:
                t = "".join(c["c"] for c in sp["chars"]).strip()
                for key in ("AAA", "BBB"):
                    if t == key and key not in ys:
                        ys[key] = sp["chars"][0]["origin"][1]
    return ys["BBB"] - ys["AAA"]


def main() -> None:
    if "--noexport" not in sys.argv:
        export()
    doc = pymupdf.open(DST)
    print(f"{'question':24s} {'with':>7s} {'without':>8s} {'advance':>8s} {'/1.2':>7s}   reading")
    for i, (label, fp, lp, epr) in enumerate(QUESTIONS):
        g_with, g_without = gap(doc[2 * i]), gap(doc[2 * i + 1])
        adv = g_with - g_without
        pt = adv / 1.2
        note = f"endParaRPr {epr / 100:g}pt" if epr else f"prev={fp:g} next={lp:g}"
        print(f"{label:24s} {g_with:7.2f} {g_without:8.2f} {adv:8.2f} {pt:7.2f}   {note}")


if __name__ == "__main__":
    main()
