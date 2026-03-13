"""
Derive Word's exact line height formula.

Key findings so far:
- docDefaults: rPrDefault has sz=22 (11pt), font=minorHAnsi (Calibri)
- w:line="240" w:lineRule="auto" = single spacing
- Word considers the paragraph mark's font formatting for minimum line height
- fsSelection USE_TYPO_METRICS flag changes which metrics are used

Approach:
1. Extract fsSelection from fonts to check USE_TYPO_METRICS
2. Compute line height at various DPIs with GDI rounding
3. Test hypothesis: max(run_font_height, paragraph_default_font_height)
"""
import json
import os
import struct

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_METRICS = os.path.join(SCRIPT_DIR, "..", "..", "crates", "oxidocs-core", "src", "font", "data", "font_metrics.json")
ISOLATED = os.path.join(SCRIPT_DIR, "output", "isolated_analysis.json")

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


def load_data():
    with open(FONT_METRICS, encoding="utf-8") as f:
        fm_list = json.load(f)
    font_metrics = {fm["family"]: fm for fm in fm_list}

    with open(ISOLATED, encoding="utf-8") as f:
        isolated = json.load(f)

    return font_metrics, isolated


def compute_gdi_line_height(fm, pt_size, dpi=96):
    """Simulate Windows GDI TEXTMETRIC calculation."""
    upm = fm["units_per_em"]
    # lfHeight = -round(pt_size * dpi / 72) → em height in pixels
    ppem = round(pt_size * dpi / 72)

    # GDI uses win metrics (unless USE_TYPO_METRICS is set)
    win_asc = fm["win_ascent"]
    win_desc = fm["win_descent"]

    # tmAscent = ceil(winAscent * ppem / UPM)
    # tmDescent = ceil(winDescent * ppem / UPM)
    import math
    tm_ascent = math.ceil(win_asc * ppem / upm)
    tm_descent = math.ceil(win_desc * ppem / upm)
    tm_height = tm_ascent + tm_descent

    # Convert back to points
    return tm_height * 72 / dpi


def compute_gdi_typo_line_height(fm, pt_size, dpi=96):
    """Simulate with typo metrics (USE_TYPO_METRICS flag)."""
    upm = fm["units_per_em"]
    ppem = round(pt_size * dpi / 72)

    typo_asc = fm["typo_ascender"]
    typo_desc = abs(fm["typo_descender"])
    typo_gap = max(0, fm["typo_line_gap"])

    import math
    tm_ascent = math.ceil(typo_asc * ppem / upm)
    tm_descent = math.ceil(typo_desc * ppem / upm)
    ext_leading = math.ceil(typo_gap * ppem / upm)

    total = tm_ascent + tm_descent + ext_leading
    return total * 72 / dpi


def main():
    font_metrics, isolated = load_data()

    print("=" * 120)
    print("HYPOTHESIS: GDI TEXTMETRIC-based line height at various DPIs")
    print("=" * 120)

    for dpi in [72, 96, 144, 150, 288]:
        print(f"\n--- DPI = {dpi} ---")
        print(f"{'Font':<12} {'Size':>5} {'Measured':>8} {'GDI_win':>8} {'err':>7} {'GDI_typo':>9} {'err':>7}")
        print("-" * 65)

        total_err_win = 0
        total_err_typo = 0
        count = 0

        for entry in isolated:
            sng = entry.get("single_nogrid")
            if not sng:
                continue
            fid = entry["font_id"]
            size = entry["size_pt"]
            fm_key = FM_MAP.get(fid)
            if not fm_key or fm_key not in font_metrics:
                continue
            fm = font_metrics[fm_key]

            gdi_win = compute_gdi_line_height(fm, size, dpi)
            gdi_typo = compute_gdi_typo_line_height(fm, size, dpi)
            err_win = sng - gdi_win
            err_typo = sng - gdi_typo

            print(f"  {fid:<12} {size:>5} {sng:>8.3f} {gdi_win:>8.3f} {err_win:>+7.3f} {gdi_typo:>9.3f} {err_typo:>+7.3f}")
            total_err_win += abs(err_win)
            total_err_typo += abs(err_typo)
            count += 1

        if count:
            print(f"  Avg |err| win: {total_err_win/count:.3f}pt  typo: {total_err_typo/count:.3f}pt")

    # Hypothesis: max(run_font_gdi_height, default_font_gdi_height)
    # Default font from docDefaults: Calibri 11pt
    print("\n\n" + "=" * 120)
    print("HYPOTHESIS: max(actual_font_height, default_Calibri_11pt_height)")
    print("Testing with various DPIs")
    print("=" * 120)

    calibri = font_metrics["Calibri"]

    for dpi in [72, 96, 144, 150]:
        default_height = compute_gdi_line_height(calibri, 11.0, dpi)
        default_height_typo = compute_gdi_typo_line_height(calibri, 11.0, dpi)

        print(f"\n--- DPI={dpi}, Calibri 11pt base: win={default_height:.3f} typo={default_height_typo:.3f} ---")
        print(f"{'Font':<12} {'Size':>5} {'Measured':>8} {'max_win':>8} {'err':>7} {'max_typo':>9} {'err':>7}")
        print("-" * 70)

        total_err = 0
        count = 0

        for entry in isolated:
            sng = entry.get("single_nogrid")
            if not sng:
                continue
            fid = entry["font_id"]
            size = entry["size_pt"]
            fm_key = FM_MAP.get(fid)
            if not fm_key or fm_key not in font_metrics:
                continue
            fm = font_metrics[fm_key]

            # max of actual font height and default font height
            gdi_win = compute_gdi_line_height(fm, size, dpi)
            gdi_typo = compute_gdi_typo_line_height(fm, size, dpi)

            max_win = max(gdi_win, default_height)
            max_typo = max(gdi_typo, default_height_typo)

            err_win = sng - max_win
            err_typo = sng - max_typo

            print(f"  {fid:<12} {size:>5} {sng:>8.3f} {max_win:>8.3f} {err_win:>+7.3f} {max_typo:>9.3f} {err_typo:>+7.3f}")
            total_err += abs(err_win)
            count += 1

        if count:
            print(f"  Avg |err| max_win: {total_err/count:.3f}pt")

    # Hypothesis: floating point (no GDI rounding), max with default
    print("\n\n" + "=" * 120)
    print("HYPOTHESIS: max(float_win, float_default) — no DPI rounding")
    print("=" * 120)

    def float_win_height(fm, size):
        return (fm["win_ascent"] + fm["win_descent"]) / fm["units_per_em"] * size

    def float_typo_height(fm, size):
        return (fm["typo_ascender"] + abs(fm["typo_descender"]) + max(0, fm["typo_line_gap"])) / fm["units_per_em"] * size

    default_win = float_win_height(calibri, 11.0)
    default_typo = float_typo_height(calibri, 11.0)

    print(f"Calibri 11pt: win={default_win:.4f} typo={default_typo:.4f}")
    print(f"{'Font':<12} {'Size':>5} {'Measured':>8} | {'font_win':>8} {'max_w':>7} {'err':>7} | {'font_typo':>9} {'max_t':>7} {'err':>7}")
    print("-" * 90)

    for entry in isolated:
        sng = entry.get("single_nogrid")
        if not sng:
            continue
        fid = entry["font_id"]
        size = entry["size_pt"]
        fm_key = FM_MAP.get(fid)
        if not fm_key or fm_key not in font_metrics:
            continue
        fm = font_metrics[fm_key]

        fw = float_win_height(fm, size)
        ft = float_typo_height(fm, size)
        mw = max(fw, default_win)
        mt = max(ft, default_typo)

        print(f"  {fid:<12} {size:>5} {sng:>8.3f} | {fw:>8.3f} {mw:>7.3f} {sng-mw:>+7.3f} | {ft:>9.3f} {mt:>7.3f} {sng-mt:>+7.3f}")


if __name__ == "__main__":
    main()
