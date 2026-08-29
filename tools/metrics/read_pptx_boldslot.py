# -*- coding: utf-8 -*-
"""Read each bold-slot arm's advances and say which metric PowerPoint used."""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx_hmtx_probe import design_width, hinted_width, load_part  # noqa: E402
from gen_pptx_boldslot import FAMILY, OUT, TEXT, source_parts  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main() -> None:
    regular, bold = source_parts()
    load_part(regular, FAMILY)
    load_part(bold, FAMILY + " Bold")
    print(f"{'arm':8s} {'PDF pen':>9} {'design':>9} {'GDI 700':>9} {'bold part':>10}   font in PDF")
    for arm in ("slot", "noslot"):
        pdf = OUT / f"{arm}.pdf"
        if not pdf.exists():
            print(f"  {arm}: not exported")
            continue
        d = pymupdf.open(pdf)
        best = None
        for b in d[0].get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                for s in l["spans"]:
                    t = "".join(c["c"] for c in s["chars"])
                    if TEXT[:10] in t:
                        best = (t, s["size"], s["chars"], s["font"])
        if not best:
            print(f"  {arm}: probe text not found")
            continue
        t, size, chars, font = best
        pen = chars[-1]["origin"][0] - chars[0]["origin"][0]
        sub = t[:-1]
        print(f"  {arm:6s} {pen:9.2f} {design_width(FAMILY, 400, sub, size):9.2f} "
              f"{hinted_width(FAMILY, 700, sub, size):9.2f} "
              f"{design_width(FAMILY + ' Bold', 400, sub, size):10.2f}   {font}")


if __name__ == "__main__":
    main()
