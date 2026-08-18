# -*- coding: utf-8 -*-
"""Export the wrap-width probe and read each line's span from PowerPoint."""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf
import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\wrapwidth\wrapwidth.pptx").resolve()
DST = SRC.with_suffix(".pdf")
ARMS = ["plain", "hang18", "hang36", "hang18_nobu", "firstind18",
        "marL36_ind0", "marL36_hang18"]
# box: off 914400 EMU = 72pt, width 3200400 EMU = 252pt; default insets 7.2/7.2
BOX_L, BOX_W, L_INS, R_INS = 72.0, 252.0, 7.2, 7.2


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


def main() -> None:
    if "--noexport" not in sys.argv:
        export()
    doc = pymupdf.open(DST)
    inner_l, inner_r = BOX_L + L_INS, BOX_L + BOX_W - R_INS
    print(f"inner area {inner_l:.1f} .. {inner_r:.1f}  (width {inner_r - inner_l:.1f}pt)")
    for i, label in enumerate(ARMS):
        rows = []
        for blk in doc[i].get_text("rawdict")["blocks"]:
            for ln in blk.get("lines", []):
                for sp in ln["spans"]:
                    t = "".join(c["c"] for c in sp["chars"]).strip()
                    if not t or sp["chars"][0]["origin"][1] < 100:
                        continue
                    ch = sp["chars"]
                    x0 = ch[0]["origin"][0]
                    last = ch[-1]
                    x1 = last["origin"][0] + (last["bbox"][2] - last["bbox"][0])
                    rows.append((ch[0]["origin"][1], x0, x1, t))
        rows.sort()
        print(f"\n{label}")
        for y, x0, x1, t in rows:
            print(f"   x {x0:7.2f}..{x1:7.2f}  right-gap {inner_r - x1:6.2f}  "
                  f"span {x1 - x0:6.2f}   {t[:46]}")


if __name__ == "__main__":
    main()
