# -*- coding: utf-8 -*-
"""Export the missing-face probe with PowerPoint, so its PDF names the faces.

NEVER run this while the renderer is producing PNGs -- a live PowerPoint COM
instance during a render run has corrupted whole decks before
(`pptx_com_render_must_not_overlap`).

Usage: python tools/metrics/export_pptx_missfont.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
PROBE = REPO / "pipeline_data" / "pptx_probes" / "missfont" / "missfont.pptx"


def main() -> None:
    if not PROBE.exists():
        sys.exit(f"{PROBE} is not there -- run gen_pptx_missfont.py first")
    out = PROBE.with_suffix(".pdf")
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(PROBE.resolve()), WithWindow=False)
        try:
            pres.SaveAs(str(out), 32)  # 32 = ppSaveAsPDF
        finally:
            pres.Close()
    finally:
        app.Quit()
    print("wrote", out)


if __name__ == "__main__":
    main()
