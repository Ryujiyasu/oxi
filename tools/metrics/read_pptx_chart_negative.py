# -*- coding: utf-8 -*-
"""Read the negative-value probe PDF: axis labels, gridlines, axis lines, bars."""
import sys
import fitz

PDF = r"pipeline_data\pptx_probes\chart_negative\chart_negative.pdf"
NAMES = ["N1 col mixed", "N2 col all-neg", "N3 line mixed", "N4 scatter Ymix",
         "N5 scatter XYmix", "N6 area mixed", "N7 bar mixed", "N8 col stacked"]


def spans(page):
    out = []
    for blk in page.get_text("rawdict")["blocks"]:
        for ln in blk.get("lines", []):
            for sp in ln.get("spans", []):
                chars = sp.get("chars", [])
                if not chars:
                    continue
                txt = "".join(c["c"] for c in chars)
                if not txt.strip():
                    continue
                x0 = min(c["origin"][0] for c in chars)
                x1 = max(c["origin"][0] + c.get("adv", 0) for c in chars)
                out.append((txt, x0, x1, sp["origin"][1], sp["size"]))
    return out


def main() -> None:
    doc = fitz.open(PDF)
    pages = [int(a) for a in sys.argv[1:]] or list(range(len(doc)))
    for pi in pages:
        page = doc[pi]
        print("=" * 78)
        print("%s  (page %d)" % (NAMES[pi] if pi < len(NAMES) else "?", pi + 1))
        hlines, vlines, rects = [], [], []
        for d in page.get_drawings():
            for it in d["items"]:
                if it[0] == "l":
                    p, q = it[1], it[2]
                    if abs(p.y - q.y) < 0.2 and abs(p.x - q.x) > 20:
                        hlines.append((round(p.y, 2), round(min(p.x, q.x), 2),
                                       round(max(p.x, q.x), 2), d.get("width")))
                    elif abs(p.x - q.x) < 0.2 and abs(p.y - q.y) > 20:
                        vlines.append((round(p.x, 2), round(min(p.y, q.y), 2),
                                       round(max(p.y, q.y), 2), d.get("width")))
                elif it[0] == "re":
                    r = it[1]
                    f = d.get("fill")
                    fh = ("#%02X%02X%02X" % tuple(int(round(c * 255)) for c in f)) if f else None
                    rects.append((round(r.x0, 2), round(r.y0, 2), round(r.x1, 2),
                                  round(r.y1, 2), fh))
        print(" H-lines (y, x0, x1, w):")
        for h in sorted(set(hlines)):
            print("   ", h)
        print(" V-lines (x, y0, y1, w):")
        for v in sorted(set(vlines)):
            print("   ", v)
        print(" Rects (x0,y0,x1,y1,fill):")
        for r in sorted(set(rects), key=lambda t: (t[0], t[1])):
            if r[4] is not None and (r[2] - r[0]) > 3 and (r[3] - r[1]) > 3:
                print("   ", r)
        print(" Spans (text, x0, x1, baseline, size):")
        for s in sorted(spans(page), key=lambda t: (round(t[3], 1), t[1])):
            print("    %-12r x %7.2f..%7.2f  bl %7.2f  sz %.2f" % s)


if __name__ == "__main__":
    main()
