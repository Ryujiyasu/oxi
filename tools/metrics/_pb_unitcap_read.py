# -*- coding: utf-8 -*-
"""Read unitcap.pdf against the ARMS of _pb_unitcap_gen.py in DOCUMENT ORDER (a
cursor walks the lines, so rotated-kana prefix collisions between arms cannot
mis-assign a line).  Usage: python _pb_unitcap_read.py [prefix] [--full]
"""
import importlib.util
import os
import sys

import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("u", os.path.join(HERE, "_pb_unitcap_gen.py"))
u = importlib.util.module_from_spec(spec)
spec.loader.exec_module(u)
want = [a for a in sys.argv[1:] if not a.startswith("--")]
sect = next((a.split("=")[1] for a in sys.argv[1:] if a.startswith("--sect=")), "4")
cs = next((a.split("=")[1] for a in sys.argv[1:] if a.startswith("--charspace=")), None)
DOCNAME = ("unitcap" if sect == "4" else "unitcap" + sect) + ("" if cs is None else "_cs" + cs)
full = "--full" in sys.argv
MARKS = "、。（）「」・1280びab"


def oxi_lines(path):
    """Oxi --dump-layout lines in document order, split into column pieces."""
    import json
    from collections import defaultdict
    pages = json.load(open(path, encoding="utf-8"))["pages"]
    out = []
    for pi, pg in enumerate(pages):
        byy = defaultdict(list)
        for el in pg.get("elements") or []:
            tx = el.get("text") or ""
            if tx.strip():
                byy[round(el.get("y", -1), 1)].append((el.get("x", -1), tx))
        for y in sorted(byy):
            items = sorted(byy[y])
            cur, prev = [], None
            for x, tx in items:
                if prev is not None and x - prev > 40:
                    out.append((pi, y, cur[0][0], "".join(a for _, a in cur)))
                    cur = []
                cur.append((x, tx))
                prev = x
            if cur:
                out.append((pi, y, cur[0][0], "".join(a for _, a in cur)))
    out.sort(key=lambda r: (r[0], 0 if r[2] < 300 else 1, r[1]))
    return out


def main():
    d = fitz.open(os.path.join(u.OUT, DOCNAME + ".pdf"))
    oxi_path = next((a.split("=", 1)[1] for a in sys.argv[1:] if a.startswith("--oxi=")), None)
    oxi = oxi_lines(oxi_path) if oxi_path else None
    ocur = 0
    lines = []
    for pno in range(len(d)):
        for b in d[pno].get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                chars = [c for s in l["spans"] for c in s["chars"] if c["c"] != " "]
                t = "".join(c["c"] for c in chars)
                if t:
                    lines.append((pno, l["bbox"][1], l["bbox"][0], t, chars))
    lines.sort(key=lambda r: (r[0], 0 if r[2] < 300 else 1, r[1]))
    cur = 0
    for label, sp, text in u.ARMS:
        body = (text[5:] if text.startswith("SPACE") else text).replace("|", "").replace("‍", "")
        if body.startswith("PPR{"):
            body = body[body.index("}") + 1:]
        if text.startswith("SPACE"):
            body = "　" + body
        n_expect = len(body) - len(u.FILL) if body.endswith(u.FILL) else len(body)
        hit = None
        for i in range(cur, len(lines)):
            if lines[i][3][:4] == body.lstrip("　")[:4]:
                hit = i
                break
        if hit is None:
            print("%-13s (not found after cursor)" % label)
            continue
        cur = hit + 1
        pno, y, x0, t, chars = lines[hit]
        if want and not any(label.startswith(w) for w in want):
            continue
        end = chars[-1]["origin"][0] + (chars[-1]["bbox"][2] - chars[-1]["bbox"][0])
        nxt_origin = chars[-1]["origin"][0]
        sel = [(j, c) for j, c in enumerate(chars) if full or c["c"] in MARKS]
        adv = " ".join("%s%.1f" % (c["c"], (chars[j + 1]["origin"][0] if j + 1 < len(chars) else c["bbox"][2]) - c["origin"][0]) for j, c in sel)
        nxt = lines[hit + 1][3][:3] if hit + 1 < len(lines) else ""
        print("%-13s line1=%2d/%2d last_org=%6.1f end=%6.1f %s | %s | next: %s" % (
            label, len(t), n_expect, nxt_origin - x0, end - x0, "GRANT" if len(t) >= n_expect else "refuse", adv, nxt))
        if oxi is not None:
            ohit = next((i for i in range(ocur, len(oxi)) if oxi[i][3][:4] == body.lstrip("　")[:4]), None)
            if ohit is not None:
                ocur = ohit + 1
                o1 = oxi[ohit]; o2 = oxi[ohit + 1] if ohit + 1 < len(oxi) else None
                ob = 51.0 if o1[2] < 300 else 308.7
                print("%-13s   OXI line1 x0=%.2f n=%d | line2 x0=%.2f n=%d %s" % ("", o1[2] - ob, len(o1[3]), (o2[2] - (51.0 if o2[2] < 300 else 308.7)) if o2 else -1, len(o2[3]) if o2 else 0, o2[3][:10] if o2 else ""))
        if "--two" in sys.argv and hit + 1 < len(lines):
            p2, y2, x2, t2, c2 = lines[hit + 1]
            base = 51.0 if x2 < 300 else 308.7
            print("%-13s   line1 x0=%.2f n=%d | line2 indent=%.2f n=%d %s" % ("", x0 - (51.0 if x0 < 300 else 308.7), len(t), x2 - base, len(t2), t2[:10]))


if __name__ == "__main__":
    main()
