"""Add real Cambria Math metrics to font_metrics_compact.json.

Same shape as S819 (Segoe UI) / S1086 (Calibri Light) / S1140 (the sweep):
pure data, no code change -- `registry.get("Cambria Math")` hits on the exact
name once the entry exists. Until now it did not, so every Cambria Math run
measured as the Calibri-class default.

WHY IT MATTERS. S880 flattens a plain inline `<m:oMath>` into a run whose
family IS "Cambria Math", and S1252 draws structured inline maths the same way,
so this face carries every inline equation in the corpus. `reference__0042471c`
p5 has one run of it -- `SIGMA Dewan Pengawas Syariah (Sharia Board Committee)` --
and Word's own per-glyph advances (`get_text("rawdict")` origin deltas out of
the reference PDF, size 9.96) put the text at **239.86pt** where Oxi painted
**258.5** (+7.8%).

★THE CODEPOINT SET IS WIDER THAN EVERY OTHER FACE'S. A maths run is not
written in ASCII: Word substitutes each letter for its MATH ITALIC codepoint
(`a` -> U+1D44E), and those are the glyphs whose advances it walks. The shared
Verdana-derived codepoint list every other entry uses stops long before
U+1D400, so an entry built the usual way would still measure the whole run at
the 0.5em fallback. This face therefore also carries:

    U+1D400-U+1D7FF   mathematical alphanumeric symbols (the italic letters)
    U+2100-U+23FF     letterlike symbols and mathematical operators (SIGMA, INTEGRAL,
                      U+210E PLANCK CONSTANT = the maths `h`)
    U+0370-U+03FF     Greek

VERIFICATION against those Word-measured advances (em, size 9.96):

    U+1D44E a  table 0.5571  Word 0.5570      U+1D45A m  0.8384  0.8380
    U+1D452 e  0.4961        0.4960           U+1D45B n  0.5737  0.5740
    U+1D456 i  0.3169        0.3130           U+1D45C o  0.5327  0.5330
    U+1D45F r  0.4756        0.4700           U+1D461 t  0.3945  0.3950
    U+1D464 w  0.7388        0.7320           U+1D446 S  0.5293  0.5290

worst |delta| = 0.0068em, and Word's figures are read off a JUSTIFIED line, so
that residual is its distributed spacing. Summed over the real run the table
predicts 238.47pt against Word's 239.86 (0.6%).

usage: python tools/metrics/add_cambria_math_metrics.py [--dry-run]
"""

import json
import os
import sys

JSON = os.path.join(
    "crates", "oxidocs-core", "src", "font", "data", "font_metrics_compact.json"
)
FAMILY = "Cambria Math"
FONT = "C:/Windows/Fonts/cambria.ttc"
REF_FAMILY = "Verdana"

# The ranges a maths run actually walks, on top of the shared set.
EXTRA_RANGES = [
    (0x0370, 0x03FF),   # Greek
    (0x2100, 0x23FF),   # letterlike symbols + mathematical operators
    (0x1D400, 0x1D7FF),  # mathematical alphanumeric symbols
]


def find_cambria_math(path):
    """Cambria Math is a separate FACE inside cambria.ttc, not a style."""
    from fontTools.ttLib import TTCollection

    ttc = TTCollection(path)
    for i, font in enumerate(ttc):
        if "MATH" not in font:
            continue
        name = font["name"].getName(1, 3, 1, 1033)
        if name and "Cambria Math" in str(name):
            return i
    raise SystemExit("Cambria Math not found in %s" % path)


def extract(path, face_num, codepoints):
    from fontTools.ttLib import TTFont

    f = TTFont(path, fontNumber=face_num, lazy=True)
    head, hhea, os2 = f["head"], f["hhea"], f["OS/2"]
    upm = head.unitsPerEm
    cmap = f.getBestCmap() or {}
    hmtx = f["hmtx"]
    widths, missing = {}, 0
    for cp in codepoints:
        g = cmap.get(int(cp))
        if g is None:
            missing += 1
            continue
        widths[str(cp)] = int(round(hmtx[g][0] * 2048 / upm))
    e = {
        "family": FAMILY,
        "units_per_em": 2048,
        "ascender": int(round(hhea.ascent * 2048 / upm)),
        "descender": int(round(hhea.descent * 2048 / upm)),
        "line_gap": int(round(hhea.lineGap * 2048 / upm)),
        "win_ascent": int(round(os2.usWinAscent * 2048 / upm)),
        "win_descent": int(round(os2.usWinDescent * 2048 / upm)),
        "typo_ascender": int(round(os2.sTypoAscender * 2048 / upm)),
        "typo_descender": int(round(os2.sTypoDescender * 2048 / upm)),
        "typo_line_gap": int(round(os2.sTypoLineGap * 2048 / upm)),
        "use_typo_metrics": bool(os2.fsSelection & (1 << 7)),
        "widths": widths,
    }
    f.close()
    return e, missing


def main():
    dry = "--dry-run" in sys.argv
    data = json.load(open(JSON, encoding="utf-8"))
    before = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    # ★The family is ALREADY in the table -- with 100 codepoints, U+0020..U+2026,
    # i.e. none of the math-italic letters a maths run is actually written in.
    # So this EXTENDS the entry rather than adding one: existing keys are left
    # exactly as they are (they are the right advances for their own
    # codepoints; the ASCII `a` at 1000du is a different glyph from the math
    # italic U+1D44E at 1141du) and only the missing ones are filled in.
    existing = next((x for x in data if x.get("family") == FAMILY), None)

    ref = [x for x in data if x.get("family") == REF_FAMILY][0]
    codepoints = set(int(k) for k in ref["widths"])
    for lo, hi in EXTRA_RANGES:
        codepoints.update(range(lo, hi + 1))
    codepoints = sorted(codepoints)

    face = find_cambria_math(FONT)
    e, missing = extract(FONT, face, codepoints)
    nat = (
        (e["typo_ascender"] + abs(e["typo_descender"]) + e["typo_line_gap"])
        if e["use_typo_metrics"]
        else max(
            (e["ascender"] + abs(e["descender"]) + e["line_gap"]),
            (e["win_ascent"] + e["win_descent"]),
        )
    ) / 2048.0
    print("%-14s face=%d hhea=%d/%d/%d win=%d/%d natural=%.6f widths=%d asked=%d"
          % (FAMILY, face, e["ascender"], e["descender"], e["line_gap"],
             e["win_ascent"], e["win_descent"], nat, len(e["widths"]), len(codepoints)))

    checks = {0x1D44E: 0.5570, 0x1D452: 0.4960, 0x1D456: 0.3130, 0x1D45A: 0.8380,
              0x1D45B: 0.5740, 0x1D45C: 0.5330, 0x1D45F: 0.4700, 0x1D461: 0.3950,
              0x1D464: 0.7320, 0x1D446: 0.5290, 0x1D437: 0.6870, 0x210E: 0.5542}
    worst = 0.0
    for cp, want in sorted(checks.items()):
        w = e["widths"].get(str(cp))
        if w is None:
            print("  U+%04X ABSENT" % cp)
            continue
        got = w / 2048.0
        worst = max(worst, abs(got - want))
        print("  U+%04X table %.4f  word %.4f  %+.4f" % (cp, got, want, got - want))
    print("  worst |delta| = %.4f em" % worst)

    if existing is not None:
        prior = dict(existing["widths"])
        added = 0
        for k, v in e["widths"].items():
            if k not in existing["widths"]:
                existing["widths"][k] = v
                added += 1
        print("  extended existing entry: %d -> %d widths (+%d)"
              % (len(prior), len(existing["widths"]), added))
        # every width the entry already carried must survive untouched
        for k, v in prior.items():
            assert existing["widths"][k] == v, ("width changed", k, v)
        # and nothing but the widths may move
        for field in ("units_per_em", "ascender", "descender", "line_gap",
                      "win_ascent", "win_descent"):
            assert existing[field] == e[field], (field, existing[field], e[field])
    else:
        data.append(e)
    if dry:
        print("(dry run, not written)")
        return 0
    out = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    if existing is None:
        kept = json.loads(out)[: len(json.loads(before))]
        assert json.dumps(kept, ensure_ascii=False, separators=(",", ":")) == before, (
            "existing entries changed!"
        )
    else:
        # additive within one entry: every OTHER entry byte-identical
        b = json.loads(before)
        a = json.loads(out)
        assert len(a) == len(b)
        for x, y in zip(a, b):
            if x.get("family") == FAMILY:
                continue
            assert json.dumps(x, ensure_ascii=False, separators=(",", ":")) ==                    json.dumps(y, ensure_ascii=False, separators=(",", ":")),                    ("entry changed", x.get("family"))
    open(JSON, "w", encoding="utf-8").write(out)
    print("wrote", JSON, "entries:", len(data))
    return 0


if __name__ == "__main__":
    sys.exit(main())
