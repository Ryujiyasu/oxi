"""Extract per-glyph tables from Cambria Math MATH table:

- MathItalicsCorrectionInfo: italic correction value per glyph (for sup positioning)
- MathTopAccentAttachment: horizontal x-offset for centering accents
- ExtendedShapeCoverage: glyphs that use extended (grown) shapes
- MathKernInfo: per-glyph cut-in values for tight sup/sub positioning
- MathVariants: vertical and horizontal glyph variants (grow chains)

Output: tools/metrics/output/cambria_math_glyph_tables.json
"""
import json, sys
from pathlib import Path
from fontTools.ttLib import TTCollection

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CAMBRIA = "C:/Windows/Fonts/cambria.ttc"
OUT = Path(__file__).with_name("output") / "cambria_math_glyph_tables.json"
OUT.parent.mkdir(parents=True, exist_ok=True)


def find_cambria_math(ttc_path):
    ttc = TTCollection(ttc_path)
    for i, font in enumerate(ttc):
        if 'MATH' in font:
            try:
                name = font['name'].getName(1, 3, 1, 1033)
                if name and "Cambria Math" in str(name):
                    return font
            except Exception:
                pass
            return font  # fallback
    return None


def resolve_glyph_name(font, glyph_id):
    """Return a human-readable name for a glyph (e.g., 'A', 'italic_A', or 'uniXXXX')."""
    glyph_names = font.getGlyphOrder()
    if 0 <= glyph_id < len(glyph_names):
        return glyph_names[glyph_id]
    return f"gid_{glyph_id}"


def get_cmap_reverse(font):
    """Build glyph_id → unicode codepoint map (first match wins)."""
    cmap_table = font['cmap']
    best = cmap_table.getBestCmap()  # dict: codepoint → glyph_name
    glyph_order = font.getGlyphOrder()
    name_to_gid = {n: i for i, n in enumerate(glyph_order)}
    rev = {}  # gid -> codepoint
    for cp, glyph_name in best.items():
        gid = name_to_gid.get(glyph_name)
        if gid is not None and gid not in rev:
            rev[gid] = cp
    return rev


def extract_italics_correction(math_table, gid_to_cp, font):
    """Return list of {gid, glyph_name, codepoint, italic_correction}."""
    ic_info = getattr(math_table.table.MathGlyphInfo, 'MathItalicsCorrectionInfo', None)
    if ic_info is None:
        return []
    coverage = ic_info.Coverage
    corrections = ic_info.ItalicsCorrection
    # coverage.glyphs is a list of glyph NAMES (not GIDs)
    glyph_names = list(coverage.glyphs) if hasattr(coverage, 'glyphs') else []
    name_to_gid = {n: i for i, n in enumerate(font.getGlyphOrder())}
    result = []
    for i, name in enumerate(glyph_names):
        gid = name_to_gid.get(name, -1)
        cp = gid_to_cp.get(gid)
        val = corrections[i].Value if i < len(corrections) else None
        result.append({"gid": gid, "glyph_name": name,
                       "codepoint": cp, "codepoint_hex": f"U+{cp:04X}" if cp else None,
                       "italic_correction": val})
    return result


def extract_top_accent_attachment(math_table, gid_to_cp, font):
    ta = getattr(math_table.table.MathGlyphInfo, 'MathTopAccentAttachment', None)
    if ta is None:
        return []
    coverage = ta.TopAccentCoverage
    attachments = ta.TopAccentAttachment
    glyph_names = list(coverage.glyphs) if hasattr(coverage, 'glyphs') else []
    name_to_gid = {n: i for i, n in enumerate(font.getGlyphOrder())}
    result = []
    for i, name in enumerate(glyph_names):
        gid = name_to_gid.get(name, -1)
        cp = gid_to_cp.get(gid)
        val = attachments[i].Value if i < len(attachments) else None
        result.append({"gid": gid, "glyph_name": name,
                       "codepoint": cp, "codepoint_hex": f"U+{cp:04X}" if cp else None,
                       "top_accent_attachment": val})
    return result


def extract_extended_shapes(math_table, gid_to_cp, font):
    cov = getattr(math_table.table.MathGlyphInfo, 'ExtendedShapeCoverage', None)
    if cov is None:
        return []
    glyph_names = list(cov.glyphs) if hasattr(cov, 'glyphs') else []
    name_to_gid = {n: i for i, n in enumerate(font.getGlyphOrder())}
    result = []
    for name in glyph_names:
        gid = name_to_gid.get(name, -1)
        cp = gid_to_cp.get(gid)
        result.append({"gid": gid, "glyph_name": name,
                       "codepoint": cp, "codepoint_hex": f"U+{cp:04X}" if cp else None})
    return result


def extract_vertical_variants(math_table, gid_to_cp, font):
    mv = getattr(math_table.table, 'MathVariants', None)
    if mv is None:
        return []
    cov = getattr(mv, 'VertGlyphCoverage', None)
    constructions = getattr(mv, 'VertGlyphConstruction', None)
    if cov is None or constructions is None:
        return []
    glyph_names = list(cov.glyphs) if hasattr(cov, 'glyphs') else []
    name_to_gid = {n: i for i, n in enumerate(font.getGlyphOrder())}
    result = []
    for i, name in enumerate(glyph_names):
        gid = name_to_gid.get(name, -1)
        cp = gid_to_cp.get(gid)
        con = constructions[i] if i < len(constructions) else None
        variants = []
        if con and hasattr(con, 'MathGlyphVariantRecord'):
            for v in con.MathGlyphVariantRecord:
                variants.append({
                    "variant_glyph": v.VariantGlyph,
                    "advance": v.AdvanceMeasurement,
                })
        result.append({"gid": gid, "glyph_name": name,
                       "codepoint": cp, "codepoint_hex": f"U+{cp:04X}" if cp else None,
                       "n_variants": len(variants), "variants": variants})
    return result


def main():
    font = find_cambria_math(CAMBRIA)
    if font is None:
        print("Cambria Math not found")
        return
    math_table = font['MATH']
    upm = font['head'].unitsPerEm
    gid_to_cp = get_cmap_reverse(font)

    print(f"UPM: {upm}")
    print(f"Total glyphs: {len(font.getGlyphOrder())}")
    print(f"Cmap entries: {len(gid_to_cp)}")

    ic = extract_italics_correction(math_table, gid_to_cp, font)
    print(f"\n=== Italic correction ({len(ic)} entries) ===")
    # Show entries with known codepoints (subset)
    named = [e for e in ic if e["codepoint"] is not None]
    print(f"  (of which {len(named)} have codepoint)")
    # Show a few examples
    for e in named[:10]:
        print(f"  {e['codepoint_hex']} {e['glyph_name']:<25}  italic_corr={e['italic_correction']}")
    if len(named) > 10:
        print(f"  ...({len(named)-10} more)")

    ta = extract_top_accent_attachment(math_table, gid_to_cp, font)
    print(f"\n=== Top accent attachment ({len(ta)} entries) ===")
    named_ta = [e for e in ta if e["codepoint"] is not None]
    print(f"  (of which {len(named_ta)} have codepoint)")
    for e in named_ta[:10]:
        print(f"  {e['codepoint_hex']} {e['glyph_name']:<25}  top_accent_attach={e['top_accent_attachment']}")

    es = extract_extended_shapes(math_table, gid_to_cp, font)
    print(f"\n=== Extended shape glyphs ({len(es)} entries) ===")
    named_es = [e for e in es if e["codepoint"] is not None]
    for e in named_es[:5]:
        print(f"  {e['codepoint_hex']} {e['glyph_name']}")

    vv = extract_vertical_variants(math_table, gid_to_cp, font)
    print(f"\n=== Vertical grow variants ({len(vv)} base glyphs) ===")
    for e in vv[:6]:
        cp_s = e['codepoint_hex'] or "?"
        print(f"  {cp_s} {e['glyph_name']:<20}  n_variants={e['n_variants']}")
        for v in e['variants'][:3]:
            print(f"    advance={v['advance']}")

    out = {
        "font": "Cambria Math", "upm": upm,
        "italic_correction": ic,
        "top_accent_attachment": ta,
        "extended_shapes": es,
        "vertical_variants": vv,
    }
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=True, indent=2)
    print(f"\nSaved → {OUT} ({OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
