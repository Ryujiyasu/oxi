"""Add real Garamond / Garamond Bold metrics to font_metrics_compact.json.

forms__002fbe2c (blindC50) sets `w:ascii="Garamond"` on 1098 runs -- the whole
body -- but the registry had no Garamond entry, so every line fell back to the
Calibri default (natural 1.2207em) where Word uses Garamond (1.125em).  That is
-0.096em per line: a 10pt line 12.207 -> 11.25, i.e. ~-1pt per line, which
accumulates over a 50-line page and spills the last one or two paragraphs.

Same shape as S819 (Segoe UI) / S855 (Arial Narrow) / S860 (Bookman) /
S866 (Aptos) / S950 (Book Antiqua) / S1001 (Open Sans) / S1007 (Comic Sans) /
S1012 (Century Gothic) / S1032 (Bell MT) / S1033 (Palatino) / S1060 (Verdana) /
S1086 (Calibri Light): pure data, no code change -- normalize/render are
passthrough and `registry.get("Garamond")` hits on the exact name.

usage: python tools/metrics/add_garamond_metrics.py [--dry-run]
"""

import json
import os
import sys

JSON = os.path.join(
    "crates", "oxidocs-core", "src", "font", "data", "font_metrics_compact.json"
)
FACES = [
    ("Garamond", r"C:\Windows\Fonts\GARA.TTF"),
    ("Garamond Bold", r"C:\Windows\Fonts\GARABD.TTF"),
]
# the codepoint set every existing entry carries
REF_FAMILY = "Verdana"


def extract(path, codepoints):
    from fontTools.ttLib import TTFont

    f = TTFont(path, fontNumber=0, lazy=True)
    head, hhea, os2 = f["head"], f["hhea"], f["OS/2"]
    upm = head.unitsPerEm
    cmap = f.getBestCmap()
    hmtx = f["hmtx"]
    widths = {}
    missing = []
    for cp in codepoints:
        g = cmap.get(int(cp))
        if g is None:
            missing.append(cp)
            continue
        widths[str(cp)] = int(round(hmtx[g][0] * 2048 / upm))
    e = {
        "family": None,
        "units_per_em": 2048,
        "ascender": int(round(hhea.ascent * 2048 / upm)),
        "descender": int(round(hhea.descent * 2048 / upm)),
        "line_gap": int(round(hhea.lineGap * 2048 / upm)),
        "win_ascent": int(round(os2.usWinAscent * 2048 / upm)),
        "win_descent": int(round(os2.usWinDescent * 2048 / upm)),
        "typo_ascender": int(round(os2.sTypoAscender * 2048 / upm)),
        "typo_descender": int(round(os2.sTypoDescender * 2048 / upm)),
        "typo_line_gap": int(round(os2.sTypoLineGap * 2048 / upm)),
        "widths": widths,
    }
    f.close()
    return e, missing


def main():
    dry = "--dry-run" in sys.argv
    data = json.load(open(JSON, encoding="utf-8"))
    before = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    ref = [x for x in data if x.get("family") == REF_FAMILY][0]
    codepoints = sorted(int(k) for k in ref["widths"])
    have = {x.get("family") for x in data}

    for fam, path in FACES:
        if fam in have:
            print("already present:", fam)
            continue
        if not os.path.exists(path):
            print("MISSING FONT FILE:", path)
            return 1
        e, missing = extract(path, codepoints)
        e["family"] = fam
        nat = max(
            (e["ascender"] + abs(e["descender"]) + e["line_gap"]),
            (e["win_ascent"] + e["win_descent"]),
        ) / 2048.0
        print(
            "%-16s upm=2048 hhea=%d/%d/%d win=%d/%d natural=%.6f widths=%d missing=%d"
            % (
                fam,
                e["ascender"],
                e["descender"],
                e["line_gap"],
                e["win_ascent"],
                e["win_descent"],
                nat,
                len(e["widths"]),
                len(missing),
            )
        )
        data.append(e)

    if dry:
        print("(dry run, not written)")
        return 0

    # additive only: every pre-existing entry must be byte-identical
    out = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    kept = json.loads(out)[: len(json.loads(before))]
    assert json.dumps(kept, ensure_ascii=False, separators=(",", ":")) == before, (
        "existing entries changed!"
    )
    open(JSON, "w", encoding="utf-8").write(out)
    print("wrote", JSON, "entries:", len(data))
    return 0


if __name__ == "__main__":
    sys.exit(main())
