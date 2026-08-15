"""Add real Tw Cen MT metrics to font_metrics_compact.json.

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

usage: python tools/metrics/add_twcen_metrics.py [--dry-run]
"""

import json
import os
import sys

JSON = os.path.join(
    "crates", "oxidocs-core", "src", "font", "data", "font_metrics_compact.json"
)
FACES = [
    ("Tw Cen MT", r"C:\Windows\Fonts\TCM_____.TTF"),
    # S1138 (2026-08-15): Trebuchet MS is installed on every Windows box and named
    # by 4 corpus docs, but had no entry, so its lines fell back to Calibri --
    # _pb_linepitch_gen.py measures Word 1.16132em against Oxi's 1.22070, i.e.
    # 0.59pt on a 10pt line. The face's own hhea gives 1.16113, one 600-DPI
    # quantum from Word's figure.
    ("Trebuchet MS", "C:/Windows/Fonts/trebuc.ttf"),
    ("Trebuchet MS Bold", "C:/Windows/Fonts/trebucbd.ttf"),
    # S1139 (2026-08-15): Meiryo UI is a separate face inside meiryo.ttc, not a
    # style of Meiryo -- its descent is 430 against Meiryo's 901, so the natural
    # height is 1.27002em, not 1.5. Without an entry the name normalised to
    # Meiryo and every line came out 1.5 x 83/64 = 1.94531em where Word draws
    # 1.27002 x 83/64 = 1.65103 (_pb_linepitch_gen.py measures 1.65119). 5 docs.
    ("Meiryo UI", "C:/Windows/Fonts/meiryo.ttc#2"),
    # S1140 (2026-08-15): the rest of the sweep -- every family the corpus
    # names in a body/header/footer run that is INSTALLED here but had no
    # entry, so Oxi measured it as Calibri. Found by diffing the corpus's
    # rFonts names against the table and the machine's font directory;
    # each is verified against Word by _pb_linepitch_gen.py after the
    # rebuild. CJK names (Yu Gothic / SimSun / MS UI Gothic) are left out --
    # they already resolve through normalize_family_name.
    ("Segoe UI Emoji", "C:/Windows/Fonts/seguiemj.ttf"),
    ("Ink Free", "C:/Windows/Fonts/Inkfree.ttf"),
    ("Franklin Gothic Book", "C:/Windows/Fonts/FRABK.TTF"),
    ("Wingdings", "C:/Windows/Fonts/wingding.ttf"),
    ("Sylfaen", "C:/Windows/Fonts/sylfaen.ttf"),
    ("Lucida Sans Unicode", "C:/Windows/Fonts/l_10646.ttf"),
    ("Jokerman", "C:/Windows/Fonts/JOKERMAN.TTF"),
    ("Impact", "C:/Windows/Fonts/impact.ttf"),
    ("Eras Bold ITC", "C:/Windows/Fonts/ERASBD.TTF"),
    ("Broadway", "C:/Windows/Fonts/BROADW.TTF"),
    ("Baskerville Old Face", "C:/Windows/Fonts/BASKVILL.TTF"),
    ("Arial Rounded MT Bold", "C:/Windows/Fonts/ARLRDBD.TTF"),
]
# the codepoint set every existing entry carries
REF_FAMILY = "Verdana"


def extract(path, codepoints):
    from fontTools.ttLib import TTFont

    # "path#N" selects a face inside a TrueType Collection (meiryo.ttc holds
    # Meiryo and Meiryo UI as separate faces, not as styles of one family).
    num = 0
    if "#" in path:
        path, _, n = path.partition("#")
        num = int(n)
    f = TTFont(path, fontNumber=num, lazy=True)
    head, hhea, os2 = f["head"], f["hhea"], f["OS/2"]
    upm = head.unitsPerEm
    cmap = f.getBestCmap()
    if cmap is None:
        # A symbol font (Wingdings) ships only the (3,0) symbol cmap, where the
        # ASCII range lives at 0xF000 + cp. getBestCmap() rejects it outright.
        sym = next((t.cmap for t in f["cmap"].tables
                    if t.platformID == 3 and t.platEncID == 0), None)
        cmap = {cp: sym[0xF000 + cp] for cp in range(0x20, 0x7F)
                if sym and (0xF000 + cp) in sym} if sym else {}
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
        if not os.path.exists(path.partition("#")[0]):
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
