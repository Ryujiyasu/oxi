"""依頼 4: Extract per-char compression spec table from existing
r17_per_char_advances_2026-05-02.json for ed025 + 7f272a.

Output: a clean (char_index, char, neighbor_prev, neighbor_next,
advance_pt, natural_width, diff_pt, mechanism, fire_rule) table
that Oxi can directly implement.
"""
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = os.path.abspath(
    "pipeline_data/r17_per_char_advances_2026-05-02.json")
OUT = os.path.abspath(
    "pipeline_data/oxi_compress_spec_table_2026-05-02.json")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")
YAKUMONO_C = set("・：；！？ー―／＼")


def classify(ch):
    if ch in YAKUMONO_A:
        return "A"
    if ch in YAKUMONO_B:
        return "B"
    if ch in YAKUMONO_C:
        return "C"
    return None


def is_cjk(ch):
    if not ch:
        return False
    o = ord(ch)
    return (0x3040 <= o <= 0x309F or 0x30A0 <= o <= 0x30FF
            or 0x4E00 <= o <= 0x9FFF or 0x3400 <= o <= 0x4DBF
            or 0xFF00 <= o <= 0xFFEF)


def is_latin_alnum(ch):
    if not ch:
        return False
    o = ord(ch)
    return (0x30 <= o <= 0x39 or 0x41 <= o <= 0x5A or 0x61 <= o <= 0x7A)


def main():
    with open(SRC, encoding="utf-8") as f:
        data = json.load(f)

    spec_tables = {}
    for label, doc in data.items():
        if "lines" not in doc:
            spec_tables[label] = {"error": doc.get("error")}
            continue
        spec_rows = []
        for line in doc["lines"]:
            advs = line["advances"]
            for i, a in enumerate(advs):
                ch = a["ch"]
                size = a["size"]
                adv = a["adv"]
                ratio = a["ratio"]
                yclass = a["yakumono_class"]
                # Compute neighbor prev/next from adjacent advances
                prev_ch = advs[i - 1]["ch"] if i > 0 else None
                next_ch = advs[i + 1]["ch"] if i + 1 < len(advs) else None
                # Determine FINAL RULE classification
                rule_match = "none"
                if yclass == "A":
                    # A compresses when preceded by A (FINAL RULE)
                    prev_class = classify(prev_ch) if prev_ch else None
                    if prev_class == "A":
                        rule_match = "A_after_A"
                elif yclass == "B":
                    # B compresses when followed by A or B
                    next_class = classify(next_ch) if next_ch else None
                    if next_class in ("A", "B"):
                        rule_match = f"B_before_{next_class}"
                # Determine actual mechanism observed
                if a.get("compressed"):
                    if ratio is not None and ratio < 0.6:
                        mech = "Mech1_half"  # compressed to ~half
                    elif ratio is not None and 0.6 <= ratio < 0.85:
                        mech = "Mech2_partial"  # partial Mech 2 distribution
                    else:
                        mech = "compressed_other"
                elif yclass and ratio is not None and 0.85 <= ratio < 0.97:
                    mech = "Mech2_slight"  # slight Mech 2 (justify-tight)
                else:
                    mech = None
                # natural width
                natural = size  # CJK natural = ppem-pt = size for non-pProportional
                # Latin alphanumeric uses different natural — for our probes
                # all chars are CJK or yakumono, so size is correct natural.
                diff = round(adv - natural, 4) if natural else None
                if (yclass is not None
                        or (ratio is not None and ratio < 0.97)):
                    spec_rows.append({
                        "i": a["i"],
                        "ch": ch,
                        "yakumono_class": yclass,
                        "prev_ch": prev_ch,
                        "next_ch": next_ch,
                        "prev_class": (classify(prev_ch)
                                        if prev_ch else None),
                        "next_class": (classify(next_ch)
                                        if next_ch else None),
                        "advance_pt": adv,
                        "natural_pt": natural,
                        "diff_pt": diff,
                        "ratio_to_natural": ratio,
                        "rule_match": rule_match,
                        "mech_observed": mech,
                        "should_oxi_compress": mech is not None,
                        "by_what": (rule_match if rule_match != "none"
                                    else "Mech2_slack"),
                    })
        spec_tables[label] = {
            "n_rows": len(spec_rows),
            "rows": spec_rows,
        }
        print(f"\n=== {label} ===")
        print(f"Total relevant chars: {len(spec_rows)}")
        # Print compressed chars only
        compressed_rows = [r for r in spec_rows if r["should_oxi_compress"]]
        print(f"Word compressed: {len(compressed_rows)}")
        print()
        print(f"{'i':>4} {'ch':>3} {'class':>5} {'prev':>3} "
              f"{'next':>3} {'adv':>6} {'nat':>6} {'diff':>7} "
              f"{'r':>5} {'mech':>14} {'rule':>14}")
        print("-" * 90)
        for r in compressed_rows:
            print(f"{r['i']:>4} {r['ch']!s:>3} {r['yakumono_class'] or '':>5} "
                  f"{r['prev_ch'] or '':>3} {r['next_ch'] or '':>3} "
                  f"{r['advance_pt']:>6.2f} {r['natural_pt']:>6.2f} "
                  f"{r['diff_pt']:>+7.2f} {r['ratio_to_natural']:>5.3f} "
                  f"{r['mech_observed'] or '':>14} {r['rule_match']:>14}")

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(spec_tables, f, ensure_ascii=False, indent=2)
    print(f"\n\nWrote spec table to {OUT}")


if __name__ == "__main__":
    main()
