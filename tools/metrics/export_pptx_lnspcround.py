# -*- coding: utf-8 -*-
"""Export the lnSpc-rounding probe to PDF with PowerPoint COM.

★One cold open, one session (`pptx_truth_pdf_first_open_is_cold`), and never
while the renderer is producing PNGs (`pptx_com_render_must_not_overlap`).
"""
from __future__ import annotations

import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\lnspcround").resolve()


def main() -> None:
    import win32com.client
    src, pdf = OUT / "lnspcround.pptx", OUT / "lnspcround.pdf"
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(src), WithWindow=False)
        try:
            pres.SaveAs(str(pdf), 32)
        finally:
            pres.Close()
    finally:
        app.Quit()
    print(f"wrote {pdf}")


if __name__ == "__main__":
    main()
