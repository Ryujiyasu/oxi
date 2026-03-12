"""
Analyze isolated line spacing tests to derive Word's exact formula.
Compares 4 modes: single_nogrid, 115_nogrid, single_grid, default
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
PDF_DIR = os.path.join(SCRIPT_DIR, "output", "pdfs_isolated")
MANIFEST = os.path.join(SCRIPT_DIR, "docx_tests_isolated", "manifest.json")
FONT_METRICS = os.path.join(SCRIPT_DIR, "..", "..", "crates", "oxidocs-core", "src", "font", "data", "font_metrics.json")

FM_MAP = {
    "yumin": "Yu Mincho Regular",
    "yugothic": "Yu Gothic Regular",
    "century": "Century",
    "tnr": "Times New Roman",
    "calibri": "Calibri",
    "arial": "Arial",
    "msmincho": "MS Mincho",
    "msgothic": "MS Gothic",
}


def extract_baseline_gap(pdf_path):
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
                lines.append({"y0": bbox[1], "y1": bbox[3], "height": bbox[3] - bbox[1]})
    doc.close()

    if len(lines) < 2:
        return None

    gaps = []
    for i in range(1, len(lines)):
        gap = lines[i]["y0"] - lines[i - 1]["y0"]
        if 0 < gap < 100:
            gaps.append(gap)

    if not gaps:
        return None

    gaps.sort()
    return gaps[len(gaps) // 2]


def main():
    with open(MANIFEST, encoding="utf-8") as f:
        manifest = json.load(f)

    with open(FONT_METRICS, encoding="utf-8") as f:
        font_metrics = {fm["family"]: fm for fm in json.load(f)}

    # Group by font+size, compare modes
    by_key = {}
    for entry in manifest:
        key = (entry["font_id"], entry["size_pt"])
        by_key.setdefault(key, {})[entry["mode"]] = entry

    print(f"{'Font':<12} {'Size':>5} | "
          f"{'single_ng':>9} {'115_ng':>8} {'single_g':>8} {'default':>8} | "
          f"{'win_h':>7} {'typo_h':>7} | "
          f"{'s_ng/win':>8} {'115/sng':>7} {'sg/sng':>7}")
    print("=" * 130)

    results = []

    for (fid, size), modes in sorted(by_key.items()):
        gaps = {}
        for mode_name, entry in modes.items():
            pdf_path = os.path.join(PDF_DIR, entry["filename"].replace(".docx", ".pdf"))
            if os.path.exists(pdf_path):
                gap = extract_baseline_gap(pdf_path)
                gaps[mode_name] = gap

        fm_key = FM_MAP.get(fid)
        win_h = typo_h = ""
        win_val = typo_val = 0
        if fm_key and fm_key in font_metrics:
            fm = font_metrics[fm_key]
            upm = fm["units_per_em"]
            win_val = (fm["win_ascent"] + fm["win_descent"]) / upm * size
            typo_val = (fm["typo_ascender"] + abs(fm["typo_descender"]) + max(0, fm["typo_line_gap"])) / upm * size
            win_h = f"{win_val:.3f}"
            typo_h = f"{typo_val:.3f}"

        sng = gaps.get("single_nogrid")
        n115 = gaps.get("115_nogrid")
        sg = gaps.get("single_grid")
        df = gaps.get("default")

        ratio_sng_win = f"{sng/win_val:.4f}" if sng and win_val else ""
        ratio_115_sng = f"{n115/sng:.4f}" if n115 and sng else ""
        ratio_sg_sng = f"{sg/sng:.4f}" if sg and sng else ""

        print(f"{fid:<12} {size:>5} | "
              f"{sng or 0:>9.3f} {n115 or 0:>8.3f} {sg or 0:>8.3f} {df or 0:>8.3f} | "
              f"{win_h:>7} {typo_h:>7} | "
              f"{ratio_sng_win:>8} {ratio_115_sng:>7} {ratio_sg_sng:>7}")

        results.append({
            "font_id": fid,
            "size_pt": size,
            "single_nogrid": round(sng, 3) if sng else None,
            "x115_nogrid": round(n115, 3) if n115 else None,
            "single_grid": round(sg, 3) if sg else None,
            "default": round(df, 3) if df else None,
            "win_height": round(win_val, 3),
            "typo_height": round(typo_val, 3),
        })

    # Summary: what formula matches single_nogrid?
    print("\n\n=== SINGLE SPACING, NO GRID (pure font line height) ===")
    print("This should reveal Word's base line height formula.\n")

    print(f"{'Font':<12} {'Size':>5} {'Measured':>8} {'win_h':>7} {'err':>7} {'typo_h':>7} {'err':>7} {'ratio':>7}")
    print("-" * 75)

    for r in results:
        sng = r["single_nogrid"]
        if not sng:
            continue
        win_err = sng - r["win_height"]
        typo_err = sng - r["typo_height"]
        ratio = sng / r["win_height"] if r["win_height"] else 0
        print(f"{r['font_id']:<12} {r['size_pt']:>5} {sng:>8.3f} "
              f"{r['win_height']:>7.3f} {win_err:>+7.3f} "
              f"{r['typo_height']:>7.3f} {typo_err:>+7.3f} "
              f"{ratio:>7.4f}")

    # Summary: does 115_nogrid = single_nogrid * 1.15?
    print("\n\n=== 1.15x MULTIPLIER CHECK ===")
    print("Is 115_nogrid exactly single_nogrid * 1.15?\n")

    print(f"{'Font':<12} {'Size':>5} {'single':>8} {'*1.15':>8} {'actual_115':>10} {'err':>7}")
    print("-" * 60)

    for r in results:
        sng = r["single_nogrid"]
        n115 = r["x115_nogrid"]
        if not sng or not n115:
            continue
        predicted = sng * 1.15
        err = n115 - predicted
        print(f"{r['font_id']:<12} {r['size_pt']:>5} {sng:>8.3f} {predicted:>8.3f} {n115:>10.3f} {err:>+7.3f}")

    # Summary: grid effect
    print("\n\n=== GRID SNAPPING EFFECT ===")
    print("Compare single_grid vs single_nogrid\n")

    print(f"{'Font':<12} {'Size':>5} {'no_grid':>8} {'grid':>8} {'diff':>7} {'grid/nogrid':>11}")
    print("-" * 60)

    for r in results:
        sng = r["single_nogrid"]
        sg = r["single_grid"]
        if not sng or not sg:
            continue
        diff = sg - sng
        ratio = sg / sng
        print(f"{r['font_id']:<12} {r['size_pt']:>5} {sng:>8.3f} {sg:>8.3f} {diff:>+7.3f} {ratio:>11.4f}")

    out = os.path.join(SCRIPT_DIR, "output", "isolated_analysis.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nResults saved to {out}")


if __name__ == "__main__":
    main()
