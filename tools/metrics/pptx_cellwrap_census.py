# -*- coding: utf-8 -*-
"""How many dev-corpus table cells hold text too wide for their column?

The renderer draws a table cell's paragraph as ONE line (main.rs table path
never calls gdi_wrap_lines), so any cell whose text is wider than its column
spills over its neighbour. This counts the exposure before the fix, straight
from the decks' XML: column width from a:tblGrid/a:gridCol (authoritative --
Google Slides exports write a placeholder graphicFrame ext), text width
estimated at 0.5 em per character, which is close enough for Latin body text
to separate "fits" from "runs 40% over".

Usage: python tools/metrics/pptx_cellwrap_census.py [--decks d25,d35]
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
PPTX_DIR = REPO / "pipeline_data" / "pptx_benchmark" / "dev" / "pptx"
EM_PER_CHAR = 0.5
DEFAULT_MARGIN_PT = 7.2  # marL + marR default (0.05in each side is 3.6pt; PPT uses 0.1in)


def cells_of(tbl: str):
    """Yield (col_index, colspan, size_pt, text) for each cell of one a:tbl."""
    for row in re.finditer(r"<a:tr\b.*?</a:tr>", tbl, re.S):
        col = 0
        for tc in re.finditer(r"<a:tc\b(.*?)</a:tc>", row.group(0), re.S):
            body = tc.group(1)
            span = int(re.search(r'gridSpan="(\d+)"', body).group(1)) if 'gridSpan="' in body else 1
            for para in re.finditer(r"<a:p>.*?</a:p>", body, re.S):
                p = para.group(0)
                text = "".join(re.findall(r"<a:t>([^<]*)</a:t>", p))
                if not text.strip():
                    continue
                sz = re.search(r'sz="(\d+)"', p)
                yield col, span, (int(sz.group(1)) / 100.0 if sz else 18.0), text
            col += span


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--decks", default=None)
    args = ap.parse_args()
    wanted = {s.strip() for s in args.decks.split(",")} if args.decks else None

    total_cells = over = 0
    per_deck: dict[str, tuple[int, int, float]] = {}
    for pptx in sorted(PPTX_DIR.glob("*.pptx")):
        did = pptx.stem.split("__")[0]
        if wanted and did not in wanted:
            continue
        z = zipfile.ZipFile(pptx)
        d_cells = d_over = 0
        d_worst = 0.0
        for name in z.namelist():
            if not re.match(r"ppt/slides/slide\d+\.xml$", name):
                continue
            xml = z.read(name).decode("utf-8", "replace")
            for tbl in re.finditer(r"<a:tbl>.*?</a:tbl>", xml, re.S):
                t = tbl.group(0)
                cols = [int(w) / 12700 for w in re.findall(r'<a:gridCol w="(\d+)"', t)]
                if not cols:
                    continue
                for col, span, sz, text in cells_of(t):
                    width = sum(cols[col:col + span]) - DEFAULT_MARGIN_PT
                    need = len(text) * EM_PER_CHAR * sz
                    d_cells += 1
                    if need > width > 0:
                        d_over += 1
                        d_worst = max(d_worst, need / width)
        if d_cells:
            per_deck[did] = (d_over, d_cells, d_worst)
            total_cells += d_cells
            over += d_over
    for did, (o, c, worst) in sorted(per_deck.items(), key=lambda kv: -kv[1][0]):
        if o:
            print(f"  {did}: {o}/{c} cells over their column, worst x{worst:.2f}")
    print(f"\n{over}/{total_cells} cells in {len(per_deck)} decks with tables")


if __name__ == "__main__":
    main()
