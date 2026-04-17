"""Extract OpenType MATH table constants from Cambria Math.

Cambria Math is the canonical font for OMML rendering. Its MATH table
contains all constants needed for proper math layout (sup/sub shifts,
fraction geometry, radical gaps, accent positioning, etc.).

Requires: fontTools (pip install fonttools)

Output:
  tools/metrics/output/cambria_math_constants.json
  tools/metrics/output/cambria_math_full_dump.json (raw table)
"""
import json
import sys
from pathlib import Path
from fontTools.ttLib import TTCollection, TTFont

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CAMBRIA_PATH = "C:/Windows/Fonts/cambria.ttc"
OUT_CONSTANTS = Path(__file__).with_name("output") / "cambria_math_constants.json"
OUT_DUMP = Path(__file__).with_name("output") / "cambria_math_full_dump.json"
OUT_CONSTANTS.parent.mkdir(parents=True, exist_ok=True)


def find_cambria_math(ttc_path):
    ttc = TTCollection(ttc_path)
    for i, font in enumerate(ttc):
        try:
            name = font['name'].getName(1, 3, 1, 1033)
            if name and "Cambria Math" in str(name):
                return font, i
        except Exception:
            continue
    # Fallback: search for MATH table existence
    for i, font in enumerate(ttc):
        if 'MATH' in font:
            return font, i
    return None, None


def extract_math_constants(font):
    """Return dict of all MathConstants values (in design units)."""
    if 'MATH' not in font:
        return {"error": "no MATH table"}
    math = font['MATH']
    table = math.table
    constants = {}

    # MathConstants object holds ~57 scalar constants + 3 tables (some minor)
    if hasattr(table, 'MathConstants'):
        mc = table.MathConstants
        # Enumerate all attributes that look like MATH constants
        for attr in dir(mc):
            if attr.startswith('_'): continue
            val = getattr(mc, attr, None)
            if val is None: continue
            # MathConstants fields: most are int or MathValueRecord
            if isinstance(val, int):
                constants[attr] = val
            elif hasattr(val, 'Value'):
                # MathValueRecord has .Value (int) + optional .DeviceTable
                constants[attr] = val.Value
    # Head/UPM — useful to scale constants to em fraction
    upm = font['head'].unitsPerEm
    return {"upm": upm, "constants": constants}


def extract_italic_correction_and_kerning(font):
    """Summary stats on GlyphInfo table (italic correction, top accent attachment)."""
    if 'MATH' not in font:
        return {}
    math = font['MATH']
    table = math.table
    summary = {}
    if hasattr(table, 'MathGlyphInfo') and table.MathGlyphInfo is not None:
        gi = table.MathGlyphInfo
        # Italic correction count
        if hasattr(gi, 'MathItalicsCorrectionInfo') and gi.MathItalicsCorrectionInfo is not None:
            ic = gi.MathItalicsCorrectionInfo
            if hasattr(ic, 'ItalicsCorrection'):
                summary["italic_correction_count"] = len(ic.ItalicsCorrection)
        # Top accent attachment
        if hasattr(gi, 'MathTopAccentAttachment') and gi.MathTopAccentAttachment is not None:
            ta = gi.MathTopAccentAttachment
            if hasattr(ta, 'TopAccentAttachment'):
                summary["top_accent_attachment_count"] = len(ta.TopAccentAttachment)
        # Extended shapes
        if hasattr(gi, 'ExtendedShapeCoverage') and gi.ExtendedShapeCoverage is not None:
            cov = gi.ExtendedShapeCoverage
            glyphs = list(cov.glyphs) if hasattr(cov, 'glyphs') else []
            summary["extended_shape_count"] = len(glyphs)
    return summary


def extract_variants_summary(font):
    """Count glyph variants and assemblies (for grow operators like ∑ ∫ ∏)."""
    if 'MATH' not in font:
        return {}
    math = font['MATH']
    table = math.table
    summary = {}
    if hasattr(table, 'MathVariants') and table.MathVariants is not None:
        mv = table.MathVariants
        if hasattr(mv, 'VertGlyphCoverage') and mv.VertGlyphCoverage:
            summary["vertical_variant_glyph_count"] = len(list(mv.VertGlyphCoverage.glyphs))
        if hasattr(mv, 'HorizGlyphCoverage') and mv.HorizGlyphCoverage:
            summary["horizontal_variant_glyph_count"] = len(list(mv.HorizGlyphCoverage.glyphs))
        if hasattr(mv, 'MinConnectorOverlap'):
            summary["MinConnectorOverlap"] = mv.MinConnectorOverlap
    return summary


def main():
    print(f"Loading {CAMBRIA_PATH}...")
    font, subfont_idx = find_cambria_math(CAMBRIA_PATH)
    if font is None:
        print("Cambria Math not found in cambria.ttc")
        return

    print(f"Found at subfont index: {subfont_idx}")
    print(f"UPM: {font['head'].unitsPerEm}")

    # Extract MathConstants
    result = extract_math_constants(font)
    upm = result["upm"]
    constants = result["constants"]
    print(f"\n=== MathConstants ({len(constants)} entries, UPM={upm}) ===")
    # Sort by common groupings
    priority_keys = [
        # Superscript
        "SuperscriptShiftUp", "SuperscriptShiftUpCramped", "SuperscriptBottomMin",
        "SuperscriptBaselineDropMax", "SuperscriptBottomMaxWithSubscript",
        # Subscript
        "SubscriptShiftDown", "SubscriptTopMax", "SubscriptBaselineDropMin",
        # Sub+sup
        "SubSuperscriptGapMin", "SuperscriptBottomMaxWithSubscript",
        "SpaceAfterScript",
        # Fraction
        "FractionNumeratorShiftUp", "FractionNumeratorDisplayStyleShiftUp",
        "FractionDenominatorShiftDown", "FractionDenominatorDisplayStyleShiftDown",
        "FractionNumeratorGapMin", "FractionNumDisplayStyleGapMin",
        "FractionDenominatorGapMin", "FractionDenomDisplayStyleGapMin",
        "FractionRuleThickness",
        # Radical
        "RadicalVerticalGap", "RadicalDisplayStyleVerticalGap",
        "RadicalRuleThickness", "RadicalExtraAscender",
        "RadicalKernBeforeDegree", "RadicalKernAfterDegree",
        "RadicalDegreeBottomRaisePercent",
        # Other geometry
        "AxisHeight", "AccentBaseHeight", "FlattenedAccentBaseHeight",
        "MathLeading", "ScriptPercentScaleDown", "ScriptScriptPercentScaleDown",
        "OverbarVerticalGap", "OverbarRuleThickness", "OverbarExtraAscender",
        "UnderbarVerticalGap", "UnderbarRuleThickness", "UnderbarExtraDescender",
        # N-ary (nary)
        "UpperLimitGapMin", "UpperLimitBaselineRiseMin",
        "LowerLimitGapMin", "LowerLimitBaselineDropMin",
        # Stack / pile
        "StackTopShiftUp", "StackTopDisplayStyleShiftUp",
        "StackBottomShiftDown", "StackBottomDisplayStyleShiftDown",
        "StackGapMin", "StackDisplayStyleGapMin",
        # Stretch stack (for under/over braces, accents on tall content)
        "StretchStackTopShiftUp", "StretchStackBottomShiftDown",
        "StretchStackGapAboveMin", "StretchStackGapBelowMin",
        # Skewed fractions
        "SkewedFractionHorizontalGap", "SkewedFractionVerticalGap",
        # Misc
        "DelimitedSubFormulaMinHeight", "DisplayOperatorMinHeight",
    ]
    for k in priority_keys:
        if k in constants:
            v = constants[k]
            frac_em = v / upm if isinstance(v, int) else None
            print(f"  {k:<50} {v:>6}  ({frac_em:.4f}em)" if frac_em is not None else f"  {k:<50} {v}")
    # Any additional constants not in priority list
    extra = {k: v for k, v in constants.items() if k not in priority_keys}
    if extra:
        print(f"\n  Other constants ({len(extra)}):")
        for k in sorted(extra):
            print(f"    {k:<50} {extra[k]}")

    # Glyph info / variants summary
    gi = extract_italic_correction_and_kerning(font)
    if gi:
        print(f"\n=== MathGlyphInfo summary ===")
        for k, v in gi.items():
            print(f"  {k:<50} {v}")
    vi = extract_variants_summary(font)
    if vi:
        print(f"\n=== MathVariants summary ===")
        for k, v in vi.items():
            print(f"  {k:<50} {v}")

    # Save JSON
    out = {
        "font": "Cambria Math",
        "path": CAMBRIA_PATH,
        "subfont_index": subfont_idx,
        "upm": upm,
        "math_constants": constants,
        "glyph_info": gi,
        "variants": vi,
    }
    with open(OUT_CONSTANTS, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=True, indent=2)
    print(f"\nSaved → {OUT_CONSTANTS}")


if __name__ == "__main__":
    main()
