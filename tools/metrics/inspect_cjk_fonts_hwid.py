"""R7.41 safety check — inspect all installed CJK fonts for hwid feature
coverage of yakumono punctuation.

R7.40 found: Meiryo has hwid (compresses), MS Mincho doesn't (no compression).
This script extends the check to MS Gothic, Yu Gothic, BIZ-UDGothic, etc.

Critical question: do the R7.38 regression docs' fonts have hwid?
- If YES, R7.41 hwid-based gate will still affect them (regression risk
  unchanged from R7.38 font-name gate)
- If NO, the hwid-gate is safer than the font-name gate

Output: per-font, list of yakumono chars with hwid halfwidth coverage.
"""
import sys
from pathlib import Path
from fontTools.ttLib import TTCollection, TTFont

# Force UTF-8 stdout on Windows so CJK glyphs print correctly.
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

FONTS_DIR = Path(r"C:\Windows\Fonts")

# Fonts to inspect (TTC + TTF). For TTCs we list each face's family name.
CJK_FONT_FILES = [
    "meiryo.ttc",
    "meiryob.ttc",
    "msgothic.ttc",
    "msmincho.ttc",
    "YuGothR.ttc",
    "YuGothB.ttc",
    "YuGothM.ttc",
    "YuGothL.ttc",
    "yumin.ttf",
    "yumindb.ttf",
    "yuminl.ttf",
    "BIZ-UDGothicR.ttc",
    "BIZ-UDGothicB.ttc",
    "BIZ-UDMinchoM.ttc",
]

YAKUMONO_CHARS = ['、', '。', '，', '．', '「', '」', '『', '』', '（', '）', '・', '：', '；']


def inspect_face(font: TTFont, family: str):
    """Return dict with hwid info for one font face."""
    try:
        gsub = font['GSUB'].table if 'GSUB' in font else None
    except Exception as e:
        return {"family": family, "error": f"GSUB load: {e}"}

    if gsub is None:
        return {"family": family, "has_gsub": False, "features": [], "hwid_chars": []}

    features = sorted({fr.FeatureTag for fr in gsub.FeatureList.FeatureRecord})
    hwid_lookup_indices = []
    for fr in gsub.FeatureList.FeatureRecord:
        if fr.FeatureTag == 'hwid':
            hwid_lookup_indices.extend(fr.Feature.LookupListIndex)

    if not hwid_lookup_indices:
        return {
            "family": family,
            "has_gsub": True,
            "features": features,
            "has_hwid": False,
            "hwid_chars": [],
        }

    cmap = font.getBestCmap()
    char_to_gid = {ch: cmap.get(ord(ch)) for ch in YAKUMONO_CHARS}

    covered = []
    for li in hwid_lookup_indices:
        lookup = gsub.LookupList.Lookup[li]
        if lookup.LookupType != 1:  # Single substitution
            continue
        for subtable in lookup.SubTable:
            mapping = getattr(subtable, 'mapping', {})
            for ch, gid in char_to_gid.items():
                if gid and gid in mapping and ch not in covered:
                    covered.append(ch)

    return {
        "family": family,
        "has_gsub": True,
        "features": features,
        "has_hwid": True,
        "hwid_chars": covered,
    }


def inspect_file(path: Path):
    """Return list of face results for a TTC/TTF file."""
    results = []
    if path.suffix.lower() == '.ttc':
        ttc = TTCollection(str(path))
        for i, face in enumerate(ttc.fonts):
            try:
                name_table = face['name']
                family = name_table.getBestFamilyName() or f"<face #{i}>"
            except Exception:
                family = f"<face #{i}>"
            res = inspect_face(face, family)
            res["file"] = path.name
            res["face_index"] = i
            results.append(res)
    else:
        face = TTFont(str(path))
        try:
            family = face['name'].getBestFamilyName() or path.stem
        except Exception:
            family = path.stem
        res = inspect_face(face, family)
        res["file"] = path.name
        res["face_index"] = 0
        results.append(res)
    return results


def main():
    all_results = []
    for fname in CJK_FONT_FILES:
        path = FONTS_DIR / fname
        if not path.exists():
            print(f"  SKIP (missing): {fname}")
            continue
        try:
            all_results.extend(inspect_file(path))
        except Exception as e:
            print(f"  ERROR on {fname}: {e}")

    print("\n=== R7.41 Safety Check: hwid coverage for yakumono ===\n")
    print(f"{'Family':<32} {'hwid?':<8} hwid-covered yakumono")
    print("-" * 80)
    for r in all_results:
        family = r["family"]
        if "error" in r:
            print(f"{family:<32} ERROR: {r['error']}")
            continue
        if not r.get("has_gsub"):
            print(f"{family:<32} no GSUB")
            continue
        has_hwid = "YES" if r["has_hwid"] else "no"
        chars = "".join(r["hwid_chars"])
        print(f"{family:<32} {has_hwid:<8} {chars}")

    print("\n=== R7.38 regression-doc font reference (from prior context) ===")
    print("  b837808: MS Gothic (10.5pt) - likely affected if MS Gothic has hwid")
    print("  ed025c, d4d126: family TBD - measure with check_word_actual_font.py")
    print("  e3c545: Meiryo (compresses, ground truth)")
    print("  0e7af1: MS Mincho (no compression, ground truth)")

    # Persist JSON for later inspection
    import json
    out_path = Path("pipeline_data/font_hwid_inspection.json")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as f:
        # FeatureList tags are short; serialize simply
        serializable = []
        for r in all_results:
            serializable.append({k: v for k, v in r.items() if k != "_raw"})
        json.dump(serializable, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {out_path}")


if __name__ == "__main__":
    main()
