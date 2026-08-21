# -*- coding: utf-8 -*-
"""Where does PowerPoint put a table cell's first baseline?

The cell path positions text by GDI's TOP (`TextOutW`), while a text frame uses
the measured first-baseline model (text_area_top + A_font x fs). This reads the
truth: for each table on a slide, the cell's top edge from the XML and the first
baseline of its text from PowerPoint's own PDF, so the offset can be compared
with what each model predicts.

Usage: python tools/metrics/read_pptx_cellbase.py <deck-prefix> <slide-number>
"""
from __future__ import annotations

import glob
import re
import sys
import zipfile
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DEV = Path(r"pipeline_data\pptx_benchmark\dev")


def main() -> None:
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    prefix, slide_no = sys.argv[1], int(sys.argv[2])
    pptx = Path(glob.glob(str(DEV / "pptx" / f"{prefix}__*.pptx"))[0])
    pdf_path = DEV / "pdf" / f"{pptx.stem}.pdf"
    z = zipfile.ZipFile(pptx)
    xml = z.read(f"ppt/slides/slide{slide_no}.xml").decode("utf-8", "replace")

    for frame in re.finditer(r"<p:graphicFrame>.*?</p:graphicFrame>", xml, re.S):
        blk = frame.group(0)
        if "<a:tbl" not in blk:
            continue
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', blk)
        rows = re.findall(r'<a:tr h="(\d+)"', blk)
        top = int(off.group(2)) / 12700 if off else 0.0
        print(f"table top = {top:.2f}pt   row heights = "
              f"{[round(int(h) / 12700, 2) for h in rows]}")
        # first row's first non-empty cell text and its marT
        tr = re.search(r"<a:tr\b.*?</a:tr>", blk, re.S).group(0)
        mar = re.search(r'marT="(\d+)"', tr)
        print(f"   marT = {int(mar.group(1)) / 12700 if mar else 'default 0.05in = 3.6'}pt")
        texts = [t for t in re.findall(r"<a:t>([^<]*)</a:t>", tr) if t.strip()]
        print(f"   first row texts: {texts[:4]}")
        break

    pdf = pymupdf.open(pdf_path)
    page = pdf[slide_no - 1]
    print("\nPowerPoint baselines (origin y of each line's first char):")
    seen = set()
    for b in page.get_text("rawdict")["blocks"]:
        if b["type"] != 0:
            continue
        for line in b["lines"]:
            for s in line["spans"]:
                ch = s["chars"]
                if not ch:
                    continue
                text = "".join(c["c"] for c in ch).strip()
                if not text or text in seen:
                    continue
                seen.add(text)
                print(f"   y={ch[0]['origin'][1]:7.2f} size={s['size']:5.2f} "
                      f"top={line['bbox'][1]:7.2f} {text[:42]!r}")


if __name__ == "__main__":
    main()
