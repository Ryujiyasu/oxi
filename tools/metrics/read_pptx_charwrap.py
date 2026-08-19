# -*- coding: utf-8 -*-
"""Export the char-wrap probe and read each line's text out of PowerPoint."""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf
import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\charwrap\charwrap.pptx").resolve()
DST = SRC.with_suffix(".pdf")
ARMS = ["latin", "hyphen", "url", "digits", "mixed", "emoji", "cjk", "giant"]
# box: left 914400 EMU = 72pt, width 2286000 EMU = 180pt, insets 7.2 each side
BOX_L, BOX_W, INS = 72.0, 180.0, 7.2


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
    inner_l, inner_r = BOX_L + INS, BOX_L + BOX_W - INS
    print(f"inner area {inner_l:.1f} .. {inner_r:.1f}  (width {inner_r - inner_l:.1f}pt)")
    for i, label in enumerate(ARMS):
        print(f"\n[{label}]")
        rows = {}
        for blk in doc[i].get_text("rawdict")["blocks"]:
            for ln in blk.get("lines", []):
                for sp in ln["spans"]:
                    t = "".join(c["c"] for c in sp["chars"])
                    if not t.strip() or sp["bbox"][1] < 60:
                        continue
                    key = round(sp["bbox"][1], 1)
                    rows.setdefault(key, []).append((sp["bbox"][0], sp["bbox"][2], t))
        for y in sorted(rows):
            parts = sorted(rows[y])
            text = "".join(p[2] for p in parts)
            x0, x1 = parts[0][0], parts[-1][1]
            print(f"  y={y:7.2f}  x {x0:6.2f}..{x1:6.2f} (right margin {inner_r - x1:5.2f})"
                  f"  {len(text):3d}ch  {text!r}")


if __name__ == "__main__":
    main()
