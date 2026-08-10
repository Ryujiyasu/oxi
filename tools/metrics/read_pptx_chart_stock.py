# -*- coding: utf-8 -*-
"""Measure the stock probe deck from the Word PDF (drawings + text)."""
import sys

import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PDF = r"pipeline_data\pptx_probes\chart_stock\chart_stock.pdf"


def main():
    only = [int(a) for a in sys.argv[1:]] or None
    doc = fitz.open(PDF)
    for pi, page in enumerate(doc, 1):
        if only and pi not in only:
            continue
        print("=" * 72)
        print("PAGE", pi)
        d = page.get_drawings()
        print("-- drawings: %d" % len(d))
        for i, p in enumerate(d):
            items = p["items"]
            kinds = {}
            for it in items:
                kinds[it[0]] = kinds.get(it[0], 0) + 1
            r = p["rect"]
            print("  [%2d] %-14s fill=%s stroke=%s w=%s rect=(%.2f,%.2f,%.2f,%.2f)"
                  % (i, str(kinds),
                     None if p["fill"] is None else tuple(round(c, 3) for c in p["fill"]),
                     None if p["color"] is None else tuple(round(c, 3) for c in p["color"]),
                     None if p["width"] is None else round(p["width"], 2),
                     r.x0, r.y0, r.x1, r.y1))
            if len(items) <= 24:
                for it in items:
                    if it[0] == "l":
                        print("        l (%.2f,%.2f)->(%.2f,%.2f)"
                              % (it[1].x, it[1].y, it[2].x, it[2].y))
                    elif it[0] == "re":
                        print("        re (%.2f,%.2f,%.2f,%.2f)"
                              % (it[1].x0, it[1].y0, it[1].x1, it[1].y1))
                    elif it[0] == "c":
                        print("        c  (%.2f,%.2f)..(%.2f,%.2f)"
                              % (it[1].x, it[1].y, it[4].x, it[4].y))
        print("-- text")
        raw = page.get_text("rawdict")
        for b in raw["blocks"]:
            for ln in b.get("lines", []):
                for s in ln["spans"]:
                    txt = "".join(c["c"] for c in s["chars"])
                    if not txt.strip():
                        continue
                    o = s["origin"]
                    print("   %-22r font=%-26s sz=%.2f origin=(%.2f,%.2f) bbox=(%.2f,%.2f,%.2f,%.2f)"
                          % (txt, s["font"], s["size"], o[0], o[1],
                             s["bbox"][0], s["bbox"][1], s["bbox"][2], s["bbox"][3]))


if __name__ == "__main__":
    main()
