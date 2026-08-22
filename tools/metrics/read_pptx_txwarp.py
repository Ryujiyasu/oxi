# -*- coding: utf-8 -*-
"""What size does PowerPoint render a prstTxWarp shape's text at?

For every `textPlain` warp shape in the dev corpus, pairs the shape's box from
the slide XML with the size PowerPoint used, read from the span in its own PDF.
Reports the size as a fraction of the box height and width so the fitting rule
is visible.

Usage: python tools/metrics/read_pptx_txwarp.py [--decks d35,d11]
"""
from __future__ import annotations

import argparse
import glob
import re
import sys
import zipfile
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--decks", default=None)
    args = ap.parse_args()
    wanted = {s.strip() for s in args.decks.split(",")} if args.decks else None

    print(f"{'deck/slide':16s} {'text':6s} {'box w x h':>16s} {'ppt size':>9s} "
          f"{'/box h':>8s} {'/box w':>8s} {'ink h':>7s}")
    for pptx in sorted((DEV / "pptx").glob("*.pptx")):
        did = pptx.stem.split("__")[0]
        if wanted and did not in wanted:
            continue
        pdfs = glob.glob(str(DEV / "pdf" / f"{pptx.stem}.pdf"))
        if not pdfs:
            continue
        z = zipfile.ZipFile(pptx)
        pdf = pymupdf.open(pdfs[0])
        for name in z.namelist():
            m = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
            if not m:
                continue
            xml = z.read(name).decode("utf-8", "replace")
            if "prstTxWarp" not in xml:
                continue
            slide_no = int(m.group(1))
            if slide_no > pdf.page_count:
                continue
            page = pdf[slide_no - 1]
            spans = []
            for b in page.get_text("rawdict")["blocks"]:
                if b["type"] != 0:
                    continue
                for line in b["lines"]:
                    for s in line["spans"]:
                        txt = "".join(c["c"] for c in s["chars"])
                        spans.append((txt.strip(), s["size"], s["bbox"]))
            for sp in re.finditer(r"<p:sp>.*?</p:sp>", xml, re.S):
                blk = sp.group(0)
                if "prstTxWarp" not in blk:
                    continue
                ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', blk)
                if not ext:
                    continue
                bw, bh = (int(v) / 12700 for v in ext.groups())
                text = "".join(re.findall(r"<a:t>([^<]*)</a:t>", blk)).strip()
                hit = next((s for s in spans if s[0] == text), None)
                if not hit:
                    print(f"{did} s{slide_no:<3d} {text[:6]!r:8s} "
                          f"{bw:7.1f} x {bh:6.1f}   (no matching span)")
                    continue
                size, bbox = hit[1], hit[2]
                ink_h = bbox[3] - bbox[1]
                print(f"{did} s{slide_no:<3d}{'':4s} {text[:6]!r:8s} "
                      f"{bw:7.1f} x {bh:6.1f} {size:9.2f} "
                      f"{size / bh:8.4f} {size / bw:8.4f} {ink_h:7.1f}")


if __name__ == "__main__":
    main()
