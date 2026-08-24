# -*- coding: utf-8 -*-
"""Read the embedded-italic-part probe back out of PowerPoint's PDF.

Each arm's line is looked up by its marker text, and the font it was set in is
resolved to a POSTSCRIPT name by extracting the embedded subset and reading name
id 6. The subset FAMILY name is an index (`43,Italic`), not a family, so it
cannot answer the question; the PostScript name can.

Also reported is the drawn advance, so a synthesised slant is visible as such:
a skewed upright part keeps the upright width, a real italic part does not.

Usage: python tools/metrics/read_pptx_italface.py
"""
from __future__ import annotations

import json
import struct
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_derive" / "italface"


def postscript_name(buf: bytes) -> str | None:
    """name id 6 out of an sfnt buffer."""
    try:
        n = struct.unpack(">H", buf[4:6])[0]
        off = None
        for i in range(n):
            rec = buf[12 + 16 * i : 12 + 16 * i + 16]
            if rec[:4] == b"name":
                off = struct.unpack(">I", rec[8:12])[0]
        if off is None:
            return None
        _, count, str_off = struct.unpack(">HHH", buf[off : off + 6])
        for i in range(count):
            pid, _, _, nid, ln, so = struct.unpack(
                ">HHHHHH", buf[off + 6 + 12 * i : off + 6 + 12 * i + 12]
            )
            if nid != 6:
                continue
            s = buf[off + str_off + so : off + str_off + so + ln]
            return s.decode("utf-16-be") if pid == 3 else s.decode("latin-1")
    except Exception:
        return None
    return None


def main() -> None:
    index = json.loads((OUT / "arms.json").read_text(encoding="utf-8"))
    rows = []
    for entry in index:
        pdf = OUT / entry["pptx"].replace(".pptx", ".pdf")
        if not pdf.exists():
            sys.exit(f"missing {pdf} -- run export_pptx_italface.py first")
        doc = pymupdf.open(pdf)
        print(f"\n=== {entry['family']}  ({pdf.name})")
        print(f"{'arm':<15} {'run_b':>5} {'run_i':>5} {'lvl_b':>5} {'lvl_i':>5} | "
              f"{'PostScript name':<30} {'width_pt':>9}")
        # subset name -> postscript, per page (resources are per-page)
        for rec in entry["arms"]:
            page = doc[rec["slide"] - 1]
            ps_by_sub = {}
            for xref, _, _, base, _, _, _ in page.get_fonts(full=True):
                try:
                    _, _, _, buf = doc.extract_font(xref)
                except Exception:
                    continue
                ps_by_sub[base.split("+")[-1]] = postscript_name(buf)
            found = None
            for blk in page.get_text("dict")["blocks"]:
                if blk.get("type") != 0:
                    continue
                for ln in blk["lines"]:
                    t = "".join(s["text"] for s in ln["spans"])
                    if rec["arm"] in t:
                        sp = ln["spans"][0]
                        found = (sp["font"], ln["bbox"][2] - ln["bbox"][0])
            if found is None:
                print(f"{rec['arm']:<15} -- line not found (outlined?) --")
                rows.append({**rec, "ps": None, "width": None})
                continue
            sub, w = found
            ps = ps_by_sub.get(sub, sub)
            def f(v):
                return "-" if v is None else str(v)
            print(f"{rec['arm']:<15} {f(rec['run_b']):>5} {f(rec['run_i']):>5} "
                  f"{f(rec['lvl_b']):>5} {f(rec['lvl_i']):>5} | {str(ps):<30} {w:9.2f}")
            rows.append({**rec, "ps": ps, "width": w})
    (OUT / "measured.json").write_text(json.dumps(rows, indent=1), encoding="utf-8")
    print(f"\nwrote {OUT / 'measured.json'}")


if __name__ == "__main__":
    main()
