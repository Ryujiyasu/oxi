# -*- coding: utf-8 -*-
"""Export the emoji-presentation probe and say which arms came out in colour."""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf
import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\emojipres\emojipres.pptx").resolve()
DST = SRC.with_suffix(".pdf")
ARMS = ["heart_plain", "heart_vs16", "hand_yes", "watch_yes", "smile_no",
        "smile_vs16", "thermo_no", "thermo_vs16", "grin_yes", "eye_no",
        "copyright_no", "letter_ctl"]


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
    for i, label in enumerate(ARMS):
        page = doc[i]
        spans = []
        for blk in page.get_text("rawdict")["blocks"]:
            for ln in blk.get("lines", []):
                for sp in ln["spans"]:
                    t = "".join(c["c"] for c in sp["chars"])
                    if sp["bbox"][1] > 60 and t.strip():
                        spans.append((t, sp["font"], round(sp["size"], 1)))
        imgs = [(round(r.width, 1), round(r.height, 1))
                for r in (page.get_image_rects(x) for x in page.get_images(full=True))
                for r in [r[0] if isinstance(r, list) else r]] if page.get_images() else []
        # A colour emoji arrives as an image and leaves no text behind.
        kind = "COLOUR" if imgs and not spans else ("text" if spans else "?")
        print(f"{label:14s} {kind:7s} spans={spans}  images={imgs}")


if __name__ == "__main__":
    main()
