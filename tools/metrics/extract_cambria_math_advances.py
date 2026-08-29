"""Extract per-glyph ADVANCE WIDTHS (hmtx) from Cambria Math.

`layout/math.rs::glyph_advance_em` guessed every math glyph's width from an
11-arm match with a 0.52 catch-all. Word's own numbers, read per-glyph out of
`reference__0042471c`'s PDF (`get_text("rawdict")` origin deltas, size 9.96),
say the guesses are far out:

    glyph        Oxi guess   Word measured
    r / U+1D45F    0.33         0.470
    t / U+1D461    0.33         0.395
    w / U+1D464    0.86         0.738
    S / U+1D446    0.68         0.529
    a / U+1D44E    0.52         0.557
    n / U+1D45B    0.52         0.574
    SIGMA U+2211   0.52         0.8795   <- the catch-all on an operator

The same document's whole maths run measures 248.6pt in Word against Oxi's
265.5 (+6.8%). Cambria Math is not in the text registry (`font_metrics_compact`
has "Cambria", a different face with different glyphs -- the maths runs on the
MATH-ITALIC codepoints U+1D44E.. which plain Cambria does not even carry).

Emit the real hmtx, keyed by CODEPOINT, in font design units plus the upm, so
the engine can divide. Only codepoints reachable through cmap are emitted.

Output: tools/metrics/output/cambria_math_advances.json
"""
import json, sys
from pathlib import Path
from fontTools.ttLib import TTCollection

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CAMBRIA = "C:/Windows/Fonts/cambria.ttc"
OUT = Path(__file__).with_name("output") / "cambria_math_advances.json"
OUT.parent.mkdir(parents=True, exist_ok=True)


def find_cambria_math(ttc_path):
    ttc = TTCollection(ttc_path)
    for font in ttc:
        if "MATH" in font:
            name = font["name"].getName(1, 3, 1, 1033)
            if name and "Cambria Math" in str(name):
                return font
    raise SystemExit("Cambria Math not found in %s" % ttc_path)


def main():
    font = find_cambria_math(CAMBRIA)
    upm = font["head"].unitsPerEm
    hmtx = font["hmtx"]
    cmap = font.getBestCmap()

    adv = {}
    for cp, gname in cmap.items():
        if gname not in hmtx.metrics:
            continue
        width = hmtx.metrics[gname][0]
        if width > 0:
            adv[str(cp)] = width

    data = {
        "font": "Cambria Math",
        "upm": upm,
        "n_glyphs": len(adv),
        "advances": adv,
    }
    OUT.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    print("wrote %s  (upm=%d, %d codepoints)" % (OUT, upm, len(adv)))

    # Cross-check against the Word-measured values above. These are read off a
    # JUSTIFIED line, so Word's spacing is distributed into the advances --
    # the MINIMUM observed occurrence of each glyph is the natural width.
    checks = {
        0x1D45F: 0.470, 0x1D461: 0.395, 0x1D464: 0.732, 0x1D446: 0.529,
        0x1D44E: 0.557, 0x1D45B: 0.574, 0x2211: 0.8795, 0x1D437: 0.687,
        0x1D452: 0.496, 0x1D456: 0.313, 0x1D45A: 0.838, 0x1D45C: 0.533,
    }
    print("\n  codepoint   table_em   word_em    delta")
    worst = 0.0
    for cp, want in sorted(checks.items()):
        got = adv.get(str(cp))
        if got is None:
            print("  U+%04X      (absent)" % cp)
            continue
        em = got / upm
        d = em - want
        worst = max(worst, abs(d))
        print("  U+%04X      %.4f     %.4f    %+.4f" % (cp, em, want, d))
    print("\n  worst |delta| = %.4f em" % worst)


if __name__ == "__main__":
    main()
