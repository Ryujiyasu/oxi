# -*- coding: utf-8 -*-
"""Export the gradient-background probe with PowerPoint (render-truth)."""
from __future__ import annotations

import sys
from pathlib import Path

import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\bg_gradient\bg_gradient.pptx").resolve()
DST = SRC.with_suffix(".pdf")


def main() -> None:
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        prs = app.Presentations.Open(str(SRC), WithWindow=False)
        try:
            print("slides:", prs.Slides.Count,
                  "size:", prs.PageSetup.SlideWidth, "x", prs.PageSetup.SlideHeight)
            prs.SaveAs(str(DST), 32)  # ppSaveAsPDF
        finally:
            prs.Close()
    finally:
        app.Quit()
    print("wrote", DST, DST.stat().st_size, "bytes")


if __name__ == "__main__":
    main()
