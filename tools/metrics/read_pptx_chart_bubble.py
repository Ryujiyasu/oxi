# -*- coding: utf-8 -*-
"""Measure the Word render of the bubble probe deck.

Bubbles are drawn as four cubic beziers.  fitz' path `rect` is the control
polygon, which is INFLATED past the circle, so the geometry is solved from the
arc END POINTS (they sit exactly on the four cardinal points).
"""
import sys

import fitz

sys.stdout.reconfigure(encoding="utf-8")

PDF = r"pipeline_data\pptx_probes\chart_bubble\chart_bubble.pdf"


def circles(page):
    out = []
    for p in page.get_drawings():
        cs = [it for it in p["items"] if it[0] == "c"]
        if len(cs) < 4:
            continue
        xs, ys = [], []
        for it in cs:
            for q in (it[1], it[4]):
                xs.append(q.x)
                ys.append(q.y)
        cx = (min(xs) + max(xs)) / 2.0
        cy = (min(ys) + max(ys)) / 2.0
        rx = (max(xs) - min(xs)) / 2.0
        ry = (max(ys) - min(ys)) / 2.0
        out.append({
            "cx": cx, "cy": cy, "rx": rx, "ry": ry,
            "fill": p.get("fill"), "stroke": p.get("color"),
            "w": p.get("width"),
        })
    return out


def lines(page):
    hs, vs = [], []
    for p in page.get_drawings():
        for it in p["items"]:
            if it[0] != "l":
                continue
            a, b = it[1], it[2]
            if abs(a.y - b.y) < 0.05 and abs(a.x - b.x) > 1:
                hs.append((round(a.y, 2), round(min(a.x, b.x), 2),
                           round(max(a.x, b.x), 2)))
            elif abs(a.x - b.x) < 0.05 and abs(a.y - b.y) > 1:
                vs.append((round(a.x, 2), round(min(a.y, b.y), 2),
                           round(max(a.y, b.y), 2)))
    return sorted(set(hs)), sorted(set(vs))


def spans(page):
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l.get("spans", []):
                t = "".join(c["c"] for c in s["chars"])
                if t.strip():
                    out.append((round(s["origin"][0], 2),
                                round(s["origin"][1], 2),
                                round(s["size"], 2), s["font"], t))
    return sorted(out, key=lambda r: (r[1], r[0]))


def main():
    doc = fitz.open(PDF)
    want = [int(a) for a in sys.argv[1:]] or list(range(1, len(doc) + 1))
    for n in want:
        pg = doc[n - 1]
        print("=" * 72)
        print("PAGE", n)
        cs = circles(pg)
        print(" circles: %d" % len(cs))
        for c in sorted(cs, key=lambda d: d["cx"]):
            print("   c=(%.2f,%.2f) rx=%.2f ry=%.2f fill=%s w=%s"
                  % (c["cx"], c["cy"], c["rx"], c["ry"], c["fill"], c["w"]))
        hs, vs = lines(pg)
        print(" H lines:", hs[:14])
        print(" V lines:", vs[:14])
        for s in spans(pg):
            print("   span", s)


if __name__ == "__main__":
    main()
