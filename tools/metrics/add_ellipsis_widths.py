# -*- coding: utf-8 -*-
"""Add the REAL U+2026 advance to every registry family whose font is on disk.

The metric tables were extracted over an ASCII-ish codepoint set, so no entry
carries U+2026. A Latin font's ellipsis then fell to FontMetrics::char_width_em's
`is_fullwidth` 1.0em heuristic — right for Times/Arial (whose ellipsis really is
2048/2048) but far too wide for Calibri (1414 = 0.690em), Segoe UI (1501),
Comic Sans (1383), Courier New (1229), Verdana (1676)...

Word's own PDF for forms__0020466f draws the dotted leader in Calibri 12pt with
an 8.28pt ellipsis advance = 1414/2048 x 12 exactly, so the per-font value is
the ground truth (a "three periods" model is falsified: it would give Calibri
1551 and Verdana 2235).

    python add_ellipsis_widths.py           # report only
    python add_ellipsis_widths.py --write   # patch font_metrics_compact.json
"""
import io
import json
import os
import sys

from fontTools.ttLib import TTFont

HERE = os.path.dirname(os.path.abspath(__file__))
JSON = os.path.join(HERE, "..", "..", "crates", "oxidocs-core", "src", "font",
                    "data", "font_metrics_compact.json")
FONTS = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")


EXTRA = os.path.join(HERE, "..", "..", "scratchpad", "fonts")


def scan(families):
    """family -> (advance, upm, file) for every family found on disk.

    `scratchpad/fonts` holds the faces S1001 downloaded for Open Sans (Word has
    no Open Sans installed here, so its PDF subset carries only the glyphs Word
    drew — and the ellipsis is not among them).
    """
    out = {}
    dirs = [FONTS] + ([EXTRA] if os.path.isdir(EXTRA) else [])
    for d in dirs:
        for fn in sorted(os.listdir(d)):
            if not fn.lower().endswith((".ttf", ".otf")):
                continue
            try:
                f = TTFont(os.path.join(d, fn), lazy=True, fontNumber=0)
                name = (f["name"].getDebugName(4) or f["name"].getDebugName(1) or "").strip()
                if name.endswith(" Regular") and name[:-8] in families:
                    name = name[:-8]
            except Exception:
                continue
            if name not in families or name in out:
                continue
            try:
                g = f.getBestCmap().get(0x2026)
                if g is None:
                    continue
                out[name] = (f["hmtx"][g][0], f["head"].unitsPerEm, fn)
            except Exception:
                continue
    return out


def main():
    data = json.load(io.open(JSON, encoding="utf-8"))
    fams = {e["family"] for e in data}
    found = scan(fams)
    write = "--write" in sys.argv
    n = 0
    for e in data:
        hit = found.get(e["family"])
        if not hit:
            continue
        adv, upm, fn = hit
        if upm != e["units_per_em"]:
            print("SKIP %-24s upm mismatch %d vs %d" % (e["family"], upm, e["units_per_em"]))
            continue
        if "8230" in e["widths"]:
            continue
        print("%-24s %5d /%d = %.5f em   (%s)" % (e["family"], adv, upm, adv / upm, fn))
        if write:
            e["widths"]["8230"] = adv
        n += 1
    print("%d families patched (of %d in the registry, %d found on disk)"
          % (n, len(fams), len(found)))
    if write:
        with io.open(JSON, "w", encoding="utf-8", newline="\n") as fh:
            json.dump(data, fh, ensure_ascii=False, separators=(",", ":"))
        print("wrote", os.path.abspath(JSON))


if __name__ == "__main__":
    main()
