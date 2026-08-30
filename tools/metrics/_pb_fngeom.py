# -*- coding: utf-8 -*-
"""Exact page-1 footnote geometry of any probe PDF, in one table.

The keep/roll models differ by a few points, so read boxes rather than derive
them from baselines: last body line bbox, rule y, note block top/bottom, note
pitch, and the two gaps that any keep test must be written in terms of --
body_bottom -> rule, and rule -> note block top.

    python tools/metrics/_pb_fngeom.py <pdf> [<pdf> ...]
"""
import os, sys
import fitz
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

print("  %-26s %7s %7s %7s %7s %6s %6s %6s %5s"
      % ("pdf", "body_bot", "rule", "note_t", "note_b", "pitch", "b->r", "r->n", "n"))
for path in sys.argv[1:]:
    doc = fitz.open(path)
    pg = doc[0]
    ph = pg.rect.height
    rules = [round(dr["rect"].y0, 2) for dr in pg.get_drawings()
             if dr["rect"].width > 50 and dr["rect"].height < 4 and dr["rect"].y0 > 300]
    rule = min(rules) if rules else None
    body, notes = [], []
    for blk in pg.get_text("dict")["blocks"]:
        for l in blk.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if not t:
                continue
            rec = (round(l["bbox"][1], 2), round(l["bbox"][3], 2))
            if rule is not None and rec[0] > rule:
                notes.append(rec)
            else:
                body.append(rec)
    body.sort()
    notes.sort()
    bb = body[-1][1] if body else float("nan")
    nt = notes[0][0] if notes else float("nan")
    nb = notes[-1][1] if notes else float("nan")
    pitch = ((notes[-1][0] - notes[0][0]) / (len(notes) - 1)) if len(notes) > 1 else float("nan")
    print("  %-26s %7.2f %7.2f %7.2f %7.2f %6.2f %6.2f %6.2f %5d"
          % (os.path.basename(path), bb, rule if rule else float("nan"), nt, nb,
             pitch, (rule - bb) if rule else float("nan"),
             (nt - rule) if rule else float("nan"), len(notes)))
    print("      page_h=%.1f  margin_bottom=%.1f  note_block_bottom_slack=%.2f"
          % (ph, ph - 72.0, (ph - 72.0) - nb))
    doc.close()
