"""
Analyze Word's default (auto) line spacing from PDF output.
Derives the formula Word uses for line height per font.
"""
import json
import os
import sys

try:
    import fitz
except ImportError:
    print("pip install pymupdf")
    sys.exit(1)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(SCRIPT_DIR, "output", "pdfs_default")
MANIFEST = os.path.join(SCRIPT_DIR, "docx_tests_default", "manifest.json")
FONT_METRICS = os.path.join(SCRIPT_DIR, "..", "..", "crates", "oxidocs-core", "src", "font", "data", "font_metrics.json")


def extract_baselines(pdf_path: str) -> list[dict]:
    """Extract line baselines and heights from PDF."""
    doc = fitz.open(pdf_path)
    lines = []

    for page in doc:
        rawdict = page.get_text("rawdict")
        for block in rawdict.get("blocks", []):
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                text = ""
                for span in line["spans"]:
                    for ci in span.get("chars", []):
                        text += ci["c"]

                if not text.strip():
                    continue

                bbox = line["bbox"]
                lines.append({
                    "text": text[:40],
                    "y0": bbox[1],
                    "y1": bbox[3],
                    "height": bbox[3] - bbox[1],
                })

    doc.close()
    return lines


def main():
    with open(MANIFEST, encoding="utf-8") as f:
        manifest = json.load(f)

    # Load font metrics for comparison
    font_metrics = {}
    if os.path.exists(FONT_METRICS):
        with open(FONT_METRICS, encoding="utf-8") as f:
            for fm in json.load(f):
                font_metrics[fm["family"]] = fm

    print(f"{'Font':<16} {'Size':>5} {'Lines':>5} {'LineH':>7} {'BaseGap':>8} "
          f"{'Gap/Size':>8} {'asc+desc':>9} {'+lgap':>9} {'Ratio':>6}")
    print("=" * 100)

    results = []

    for entry in manifest:
        pdf_path = os.path.join(PDF_DIR, entry["filename"].replace(".docx", ".pdf"))
        if not os.path.exists(pdf_path):
            continue

        lines = extract_baselines(pdf_path)
        if len(lines) < 2:
            continue

        # Compute baseline gaps between consecutive lines
        gaps = []
        for i in range(1, len(lines)):
            gap = lines[i]["y0"] - lines[i - 1]["y0"]
            if gap > 0 and gap < 100:  # sanity check
                gaps.append(gap)

        if not gaps:
            continue

        # Use median to avoid outliers from paragraph boundaries
        gaps.sort()
        median_gap = gaps[len(gaps) // 2]
        avg_height = sum(l["height"] for l in lines) / len(lines)
        size = entry["size_pt"]

        # Font metrics comparison
        fm_key = None
        font_id = entry["font_id"]
        fm_map = {
            "yumin": "Yu Mincho Regular",
            "yugothic": "Yu Gothic Regular",
            "century": "Century",
            "tnr": "Times New Roman",
            "calibri": "Calibri",
            "arial": "Arial",
            "msmincho": "MS Mincho",
            "msgothic": "MS Gothic",
        }
        fm_key = fm_map.get(font_id)

        asc_desc = ""
        asc_desc_lgap = ""
        ratio = ""
        if fm_key and fm_key in font_metrics:
            fm = font_metrics[fm_key]
            upm = fm["units_per_em"]
            asc = fm["ascender"]
            desc = abs(fm["descender"])
            lgap = max(0, fm["line_gap"])

            ad_norm = (asc + desc) / upm
            adl_norm = (asc + desc + lgap) / upm

            ad_pt = ad_norm * size
            adl_pt = adl_norm * size
            asc_desc = f"{ad_pt:.2f}"
            asc_desc_lgap = f"{adl_pt:.2f}"
            ratio = f"{median_gap / size:.3f}"

        print(f"{entry['font']:<16} {size:>5} {len(lines):>5} {avg_height:>7.2f} {median_gap:>8.2f} "
              f"{median_gap/size:>8.3f} {asc_desc:>9} {asc_desc_lgap:>9} {ratio:>6}")

        results.append({
            "font": entry["font"],
            "font_id": font_id,
            "size_pt": size,
            "line_count": len(lines),
            "avg_line_height": round(avg_height, 3),
            "median_baseline_gap": round(median_gap, 3),
            "gap_over_size": round(median_gap / size, 4),
        })

    # Write results
    out = os.path.join(SCRIPT_DIR, "output", "default_spacing_analysis.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    # Derive formula
    print("\n" + "=" * 80)
    print("FORMULA DERIVATION: gap/size ratio by font family")
    print("=" * 80)

    by_font = {}
    for r in results:
        fid = r["font_id"]
        by_font.setdefault(fid, []).append(r)

    for fid, entries in sorted(by_font.items()):
        ratios = [e["gap_over_size"] for e in entries]
        avg_ratio = sum(ratios) / len(ratios)
        fm_key = fm_map.get(fid)
        metrics_info = ""
        if fm_key and fm_key in font_metrics:
            fm = font_metrics[fm_key]
            upm = fm["units_per_em"]
            asc = fm["ascender"]
            desc = abs(fm["descender"])
            lgap = max(0, fm["line_gap"])
            adl = (asc + desc + lgap) / upm
            metrics_info = f"  (a+d+g)/upm={adl:.4f}"

        print(f"  {fid:<12}: avg ratio = {avg_ratio:.4f}  "
              f"per-size = {[f'{r:.4f}' for r in ratios]}{metrics_info}")


if __name__ == "__main__":
    main()
