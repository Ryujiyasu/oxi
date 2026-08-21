# -*- coding: utf-8 -*-
"""Re-save a .pptx through PowerPoint with its fonts EMBEDDED.

`Presentation.SaveAs` takes EmbedTrueTypeFonts as its third argument, which is
the only reliable way to produce the `p:embeddedFontLst` + EOT `.fntdata` parts
a real deck carries — they cannot be written by hand (MicroType Express).

Usage: python tools/metrics/embed_pptx_fonts.py <deck.pptx> [out.pptx]
"""
from __future__ import annotations

import os
import sys

import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PP_SAVE_AS_DEFAULT = 11  # ppSaveAsOpenXMLPresentation
MSO_TRUE = -1


def main() -> None:
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = os.path.abspath(sys.argv[1])
    dst = os.path.abspath(
        sys.argv[2] if len(sys.argv) > 2 else src.replace(".pptx", "_embedded.pptx")
    )
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(src, WithWindow=False)
    try:
        pres.SaveAs(dst, PP_SAVE_AS_DEFAULT, MSO_TRUE)
        print(f"wrote {dst}")
    finally:
        pres.Close()


if __name__ == "__main__":
    main()
