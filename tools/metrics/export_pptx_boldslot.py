# -*- coding: utf-8 -*-
"""Export both arms of the bold-slot probe to PDF with PowerPoint COM.

★Never while the renderer is producing PNGs (`pptx_com_render_must_not_overlap`),
and one cold open each (`pptx_truth_pdf_first_open_is_cold`).
"""
from __future__ import annotations

import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\boldslot").resolve()


def main() -> None:
    import win32com.client
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        for arm in ("slot", "noslot"):
            pres = app.Presentations.Open(str(OUT / f"{arm}.pptx"), WithWindow=False)
            try:
                pres.SaveAs(str(OUT / f"{arm}.pdf"), 32)
            finally:
                pres.Close()
            print(f"  {arm} -> {arm}.pdf", flush=True)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
