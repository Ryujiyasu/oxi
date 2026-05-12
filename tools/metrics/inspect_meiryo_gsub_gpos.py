"""Inspect Meiryo TTC's GSUB/GPOS tables to identify yakumono compression rules.

Strategy: read meiryo.ttc → enumerate fonts → list GSUB/GPOS features →
focus on chws (contextual half-width spacing), cjct, vert, palt, halt,
hwid, fwid, twid, qwid features that may drive CJK punctuation
compression in Word.

If chws/palt rules are found and contain "、「" and similar pair lookups,
that's Word's mechanism for pair compression.
"""
from __future__ import annotations
import sys
from fontTools.ttLib import TTCollection, TTFont
from fontTools.ttLib.tables.otBase import BaseTable

MEIRYO_TTC = r"C:\Windows\Fonts\meiryo.ttc"
MSMINCHO_TTC = r"C:\Windows\Fonts\msmincho.ttc"


def inspect_font(font: TTFont, label: str):
    print(f"\n=== {label} ===")
    name_table = font['name']
    family_name = None
    for record in name_table.names:
        if record.nameID == 1:
            try:
                family_name = record.toUnicode()
                break
            except Exception:
                pass
    print(f"Family: {family_name!r}")

    # GSUB features
    if 'GSUB' in font:
        gsub = font['GSUB'].table
        features = set()
        if gsub.FeatureList:
            for fr in gsub.FeatureList.FeatureRecord:
                features.add(fr.FeatureTag)
        print(f"GSUB features: {sorted(features)}")
    else:
        print("GSUB: not present")

    # GPOS features
    if 'GPOS' in font:
        gpos = font['GPOS'].table
        features = set()
        if gpos.FeatureList:
            for fr in gpos.FeatureList.FeatureRecord:
                features.add(fr.FeatureTag)
        print(f"GPOS features: {sorted(features)}")
    else:
        print("GPOS: not present")

    # Focus on chws lookups specifically
    target_features = ['chws', 'palt', 'pwid', 'hwid', 'qwid', 'twid', 'fwid', 'halt']
    if 'GPOS' in font:
        gpos = font['GPOS'].table
        for fr in gpos.FeatureList.FeatureRecord:
            if fr.FeatureTag in target_features:
                feature = fr.Feature
                lookup_indices = feature.LookupListIndex
                print(f"\n  GPOS '{fr.FeatureTag}' uses lookups: {lookup_indices}")
                for li in lookup_indices[:3]:  # show first 3 lookups
                    lookup = gpos.LookupList.Lookup[li]
                    print(f"    Lookup[{li}] type={lookup.LookupType}")
    if 'GSUB' in font:
        gsub = font['GSUB'].table
        for fr in gsub.FeatureList.FeatureRecord:
            if fr.FeatureTag in target_features:
                feature = fr.Feature
                lookup_indices = feature.LookupListIndex
                print(f"\n  GSUB '{fr.FeatureTag}' uses lookups: {lookup_indices}")
                for li in lookup_indices[:3]:
                    lookup = gsub.LookupList.Lookup[li]
                    print(f"    Lookup[{li}] type={lookup.LookupType}")


def main():
    for path, label in [(MEIRYO_TTC, "Meiryo"), (MSMINCHO_TTC, "MS Mincho")]:
        print(f"\n##### {path} #####")
        try:
            ttc = TTCollection(path)
            for i, font in enumerate(ttc.fonts):
                inspect_font(font, f"{label} (font #{i})")
        except Exception as e:
            print(f"Failed: {e}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
