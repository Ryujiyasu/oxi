# -*- coding: utf-8 -*-
"""Measure `_pb_fnarea`'s constants: note pitch, rule y, body pitch, last body line.

`_pb_fnkeep` pinned the A4/Calibri geometry to a full-reservation keep test with
sep ~= 13.4 and no roll band at all. `_pb_fnarea` (Letter/TNR, 8pt notes) shows
a genuine roll -- R07 stays on page 1 while NOTE7 renders on page 2 -- so read
its numbers rather than assume them, and check whether the rolled note would
have fitted geometrically had Word placed it.

    python tools/metrics/_pb_fnarea_probe.py [fna_00200 ...]
"""
import os, sys, glob
import fitz
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                   "pipeline_data", "_pb_fnarea")
MARGIN_BOTTOM = 792.0 - 72.0

names = sys.argv[1:] or [os.path.basename(p)[:-4]
                         for p in sorted(glob.glob(os.path.join(DIR, "fna_*.pdf")))]

for name in names:
    doc = fitz.open(os.path.join(DIR, name + ".pdf"))
    print("=== %s ===" % name)
    for pno in range(min(2, doc.page_count)):
        pg = doc[pno]
        rules = [round(dr["rect"].y0, 2) for dr in pg.get_drawings()
                 if dr["rect"].width > 50 and dr["rect"].height < 3 and dr["rect"].y0 > 300]
        body, notes = [], []
        for blk in pg.get_text("dict")["blocks"]:
            for l in blk.get("lines", []):
                t = "".join(s["text"] for s in l["spans"]).strip()
                if not t:
                    continue
                rec = (round(l["bbox"][1], 2), round(l["bbox"][3], 2), t[:34])
                (notes if "NOTE" in t and "ref text" in t else body).append(rec)
        body.sort()
        notes.sort()
        rule = min(rules) if rules else None
        print("  page %d  rule=%s" % (pno + 1, rule))
        for y0, y1, t in body[-3:]:
            print("    body  top=%7.2f bot=%7.2f  %s" % (y0, y1, t))
        for y0, y1, t in notes:
            print("    note  top=%7.2f bot=%7.2f  %s" % (y0, y1, t))
        if notes and len(notes) > 1:
            pitch = (notes[-1][0] - notes[0][0]) / (len(notes) - 1)
            print("    note pitch = %.3f   block %.2f..%.2f   margin_bottom %.1f"
                  % (pitch, notes[0][0], notes[-1][1], MARGIN_BOTTOM))
            if pno == 0 and body:
                nxt = notes[-1][1] + pitch
                print("    one MORE note would end at %.2f  (%s the %0.1f margin)"
                      % (nxt, "past" if nxt > MARGIN_BOTTOM else "inside", MARGIN_BOTTOM))
        if len(body) > 1:
            bp = [body[i + 1][0] - body[i][0] for i in range(len(body) - 1)]
            bp = [p for p in bp if p > 1]
            if bp:
                print("    body pitch = %.3f (min %.2f max %.2f)"
                      % (sum(bp) / len(bp), min(bp), max(bp)))
    doc.close()
