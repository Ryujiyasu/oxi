"""Compare current Oxi line height formula against Word COM measurements."""
import json, os, math

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
COM_DATA = os.path.join(SCRIPT_DIR, "output", "word_com_positions.json")
PDF_DATA = os.path.join(SCRIPT_DIR, "output", "isolated_analysis.json")
FONT_METRICS = os.path.join(SCRIPT_DIR, "..", "..", "crates", "oxidocs-core", "src", "font", "data", "font_metrics.json")

FM_MAP = {
    "yumin": "Yu Mincho Regular", "yugothic": "Yu Gothic Regular",
    "century": "Century", "tnr": "Times New Roman",
    "calibri": "Calibri", "arial": "Arial",
    "msmincho": "MS Mincho", "msgothic": "MS Gothic",
}


def load_all():
    with open(FONT_METRICS, encoding="utf-8") as f:
        fm_list = json.load(f)
    fm = {x["family"]: x for x in fm_list}

    with open(COM_DATA, encoding="utf-8") as f:
        com = json.load(f)

    with open(PDF_DATA, encoding="utf-8") as f:
        pdf = json.load(f)

    return fm, com, pdf


def oxi_line_height_ratio(fm):
    """Current Oxi: max(win, hhea) / UPM"""
    upm = fm["units_per_em"]
    win = (fm["win_ascent"] + fm["win_descent"]) / upm
    hhea_asc = fm.get("ascender", fm["win_ascent"])
    hhea_desc = abs(fm.get("descender", -fm["win_descent"]))
    hhea_gap = max(0, fm.get("line_gap", 0))
    hhea = (hhea_asc + hhea_desc + hhea_gap) / upm
    return max(win, hhea)


def word_line_height(fm, pt_size, dpi):
    """Word line height simulation at given DPI."""
    upm = fm["units_per_em"]
    ppem = round(pt_size * dpi / 72)
    tm_asc = math.ceil(fm["win_ascent"] * ppem / upm)
    tm_desc = math.ceil(fm["win_descent"] * ppem / upm)
    return (tm_asc + tm_desc) * 72 / dpi


def main():
    fm, com, pdf = load_all()

    # Current Oxi default: Calibri 11pt
    calibri = fm["Calibri"]
    oxi_default = oxi_line_height_ratio(calibri) * 11.0

    print("=" * 100)
    print("CURRENT OXI vs WORD COM (single spacing, no grid)")
    print(f"Oxi default base (Calibri 11pt): {oxi_default:.4f}pt")
    print("=" * 100)
    print(f"{'Font':<12} {'Size':>5} | {'COM':>7} {'Oxi':>7} {'err':>7} {'%err':>6} | {'PDF':>7} {'Oxi':>7} {'err':>7} {'%err':>6}")
    print("-" * 90)

    com_errors = []
    pdf_errors = []

    for c in com:
        fid = c["font_id"]
        size = c["size"]
        com_gap = c["gap"]

        fm_key = FM_MAP.get(fid)
        if not fm_key or fm_key not in fm:
            continue

        ratio = oxi_line_height_ratio(fm[fm_key])
        oxi_val = max(ratio * size, oxi_default)
        com_err = oxi_val - com_gap
        com_pct = abs(com_err) / com_gap * 100
        com_errors.append((fid, size, com_err, com_pct))

        # Find matching PDF entry
        pdf_entry = next((p for p in pdf if p["font_id"] == fid and p["size_pt"] == size), None)
        pdf_gap = pdf_entry["single_nogrid"] if pdf_entry else None
        if pdf_gap:
            pdf_err = oxi_val - pdf_gap
            pdf_pct = abs(pdf_err) / pdf_gap * 100
            pdf_errors.append((fid, size, pdf_err, pdf_pct))
            print(f"  {fid:<12} {size:>5} | {com_gap:>7.2f} {oxi_val:>7.2f} {com_err:>+7.2f} {com_pct:>5.1f}% | {pdf_gap:>7.2f} {oxi_val:>7.2f} {pdf_err:>+7.2f} {pdf_pct:>5.1f}%")
        else:
            print(f"  {fid:<12} {size:>5} | {com_gap:>7.2f} {oxi_val:>7.2f} {com_err:>+7.2f} {com_pct:>5.1f}% |")

    avg_com = sum(abs(e[2]) for e in com_errors) / len(com_errors)
    max_com = max(abs(e[2]) for e in com_errors)
    avg_com_pct = sum(e[3] for e in com_errors) / len(com_errors)

    print(f"\nCOM:  avg |err| = {avg_com:.2f}pt ({avg_com_pct:.1f}%), max |err| = {max_com:.2f}pt")

    if pdf_errors:
        avg_pdf = sum(abs(e[2]) for e in pdf_errors) / len(pdf_errors)
        max_pdf = max(abs(e[2]) for e in pdf_errors)
        avg_pdf_pct = sum(e[3] for e in pdf_errors) / len(pdf_errors)
        print(f"PDF:  avg |err| = {avg_pdf:.2f}pt ({avg_pdf_pct:.1f}%), max |err| = {max_pdf:.2f}pt")

    within_1pt = sum(1 for e in com_errors if abs(e[2]) <= 1.0)
    within_05pt = sum(1 for e in com_errors if abs(e[2]) <= 0.5)
    total = len(com_errors)
    print(f"\nCOM accuracy: {within_05pt}/{total} within 0.5pt, {within_1pt}/{total} within 1.0pt")

    # ====== Try Word hypothesis at various DPIs ======
    print("\n\n" + "=" * 100)
    print("WORD HYPOTHESIS: max(word(run_font, size, DPI), word(default_font, 11, DPI))")
    print("=" * 100)

    best_dpi = None
    best_avg = 999

    for dpi in [72, 96, 120, 144, 150, 192, 288]:
        word_default = word_line_height(calibri, 11.0, dpi)
        errors = []

        for c in com:
            fid = c["font_id"]
            size = c["size"]
            com_gap = c["gap"]
            fm_key = FM_MAP.get(fid)
            if not fm_key or fm_key not in fm:
                continue

            word_val = word_line_height(fm[fm_key], size, dpi)
            predicted = max(word_val, word_default)
            err = predicted - com_gap
            errors.append(abs(err))

        avg_err = sum(errors) / len(errors)
        max_err = max(errors)
        within = sum(1 for e in errors if e <= 0.5)

        marker = ""
        if avg_err < best_avg:
            best_avg = avg_err
            best_dpi = dpi
            marker = " <-- BEST"

        print(f"  DPI={dpi:>4}: avg |err|={avg_err:.3f}pt, max={max_err:.3f}pt, {within}/{len(errors)} within 0.5pt{marker}")

    # Show best DPI detail
    print(f"\n--- Best DPI = {best_dpi}, detailed ---")
    print(f"{'Font':<12} {'Size':>5} {'COM':>7} {'Word':>7} {'err':>7}")
    print("-" * 50)

    word_default = word_line_height(calibri, 11.0, best_dpi)

    for c in com:
        fid = c["font_id"]
        size = c["size"]
        com_gap = c["gap"]
        fm_key = FM_MAP.get(fid)
        if not fm_key or fm_key not in fm:
            continue
        word_val = word_line_height(fm[fm_key], size, best_dpi)
        predicted = max(word_val, word_default)
        err = predicted - com_gap
        print(f"  {fid:<12} {size:>5} {com_gap:>7.2f} {predicted:>7.2f} {err:>+7.2f}")

    # ====== Try with Cambria as default (python-docx template) ======
    cambria = fm.get("Cambria", calibri)
    msmincho = fm.get("MS Mincho", calibri)

    print("\n\n" + "=" * 100)
    print("HYPOTHESIS: default font = Cambria (python-docx template theme)")
    print("=" * 100)

    for dpi in [72, 96, 144, 150]:
        word_default_cambria = word_line_height(cambria, 11.0, dpi)
        errors = []

        for c in com:
            fid = c["font_id"]
            size = c["size"]
            com_gap = c["gap"]
            fm_key = FM_MAP.get(fid)
            if not fm_key or fm_key not in fm:
                continue
            word_val = word_line_height(fm[fm_key], size, dpi)
            predicted = max(word_val, word_default_cambria)
            errors.append(abs(predicted - com_gap))

        avg_err = sum(errors) / len(errors)
        max_err = max(errors)
        within = sum(1 for e in errors if e <= 0.5)
        print(f"  DPI={dpi:>4}: default={word_default_cambria:.3f}pt, avg |err|={avg_err:.3f}pt, max={max_err:.3f}pt, {within}/{len(errors)} within 0.5pt")

    # ====== Try mixed: East Asian default font differs ======
    ea_fonts = {"msgothic", "msmincho", "yugothic", "yumin"}

    print("\n\n" + "=" * 100)
    print("HYPOTHESIS: split default - Cambria 11pt (Latin) + MS Mincho 11pt (EA)")
    print("=" * 100)

    for dpi in [72, 96, 144, 150]:
        word_latin_default = word_line_height(cambria, 11.0, dpi)
        word_ea_default = word_line_height(msmincho, 11.0, dpi)
        errors = []

        for c in com:
            fid = c["font_id"]
            size = c["size"]
            com_gap = c["gap"]
            fm_key = FM_MAP.get(fid)
            if not fm_key or fm_key not in fm:
                continue
            word_val = word_line_height(fm[fm_key], size, dpi)
            if fid in ea_fonts:
                predicted = max(word_val, word_ea_default)
            else:
                predicted = max(word_val, word_latin_default)
            errors.append(abs(predicted - com_gap))

        avg_err = sum(errors) / len(errors)
        max_err = max(errors)
        within = sum(1 for e in errors if e <= 0.5)
        print(f"  DPI={dpi:>4}: latin_def={word_latin_default:.3f}, ea_def={word_ea_default:.3f}, avg |err|={avg_err:.3f}pt, max={max_err:.3f}pt, {within}/{len(errors)} within 0.5pt")

    # ====== Key insight: Word uses BOTH hAnsi + eastAsia default fonts ======
    # The paragraph mark has both Latin and EA font components.
    # Line height minimum = max(latin_default_height, ea_default_height)
    # For each RUN, the line height = max(run_font_height, paragraph_mark_max)
    print("\n\n" + "=" * 100)
    print("HYPOTHESIS: paragraph mark uses max(Cambria 11pt, MS Mincho 11pt) as minimum")
    print("i.e., ALL fonts use same default = max of both theme fonts")
    print("=" * 100)

    for dpi in [72, 96, 144, 150]:
        word_cambria = word_line_height(cambria, 11.0, dpi)
        word_msmincho = word_line_height(msmincho, 11.0, dpi)
        combined_default = max(word_cambria, word_msmincho)
        errors = []

        for c in com:
            fid = c["font_id"]
            size = c["size"]
            com_gap = c["gap"]
            fm_key = FM_MAP.get(fid)
            if not fm_key or fm_key not in fm:
                continue
            word_val = word_line_height(fm[fm_key], size, dpi)
            predicted = max(word_val, combined_default)
            errors.append(abs(predicted - com_gap))

        avg_err = sum(errors) / len(errors)
        max_err = max(errors)
        within = sum(1 for e in errors if e <= 0.5)
        print(f"  DPI={dpi:>4}: cambria={word_cambria:.3f}, msmincho={word_msmincho:.3f}, combined={combined_default:.3f}, avg |err|={avg_err:.3f}pt, max={max_err:.3f}pt, {within}/{len(errors)} within 0.5pt")

    # ====== Try hhea metrics instead of win ======
    print("\n\n" + "=" * 100)
    print("HYPOTHESIS: Word uses hhea metrics (not win) for line height calculation")
    print("=" * 100)

    def word_hhea_line_height(fm_data, pt_size, dpi):
        upm = fm_data["units_per_em"]
        ppem = round(pt_size * dpi / 72)
        asc = fm_data.get("ascender", fm_data["win_ascent"])
        desc = abs(fm_data.get("descender", -fm_data["win_descent"]))
        gap = max(0, fm_data.get("line_gap", 0))
        tm_asc = math.ceil(asc * ppem / upm)
        tm_desc = math.ceil(desc * ppem / upm)
        tm_gap = math.ceil(gap * ppem / upm)
        return (tm_asc + tm_desc + tm_gap) * 72 / dpi

    for dpi in [72, 96, 144, 150]:
        hhea_cambria = word_hhea_line_height(cambria, 11.0, dpi)
        errors = []

        for c in com:
            fid = c["font_id"]
            size = c["size"]
            com_gap = c["gap"]
            fm_key = FM_MAP.get(fid)
            if not fm_key or fm_key not in fm:
                continue
            hhea_val = word_hhea_line_height(fm[fm_key], size, dpi)
            predicted = max(hhea_val, hhea_cambria)
            errors.append(abs(predicted - com_gap))

        avg_err = sum(errors) / len(errors)
        max_err = max(errors)
        within = sum(1 for e in errors if e <= 0.5)
        print(f"  DPI={dpi:>4}: hhea default={hhea_cambria:.3f}, avg |err|={avg_err:.3f}pt, max={max_err:.3f}pt, {within}/{len(errors)} within 0.5pt")

    # ====== Try max(win, hhea) Word ======
    print("\n\n" + "=" * 100)
    print("HYPOTHESIS: max(word_win, word_hhea) per font + max with default")
    print("=" * 100)

    def word_max_line_height(fm_data, pt_size, dpi):
        return max(word_line_height(fm_data, pt_size, dpi),
                   word_hhea_line_height(fm_data, pt_size, dpi))

    for dpi in [72, 96, 144, 150]:
        default_h = word_max_line_height(cambria, 11.0, dpi)
        errors = []
        details = []

        for c in com:
            fid = c["font_id"]
            size = c["size"]
            com_gap = c["gap"]
            fm_key = FM_MAP.get(fid)
            if not fm_key or fm_key not in fm:
                continue
            run_h = word_max_line_height(fm[fm_key], size, dpi)
            predicted = max(run_h, default_h)
            err = predicted - com_gap
            errors.append(abs(err))
            details.append((fid, size, com_gap, predicted, err))

        avg_err = sum(errors) / len(errors)
        max_err = max(errors)
        within = sum(1 for e in errors if e <= 0.5)
        print(f"  DPI={dpi:>4}: default={default_h:.3f}, avg |err|={avg_err:.3f}pt, max={max_err:.3f}pt, {within}/{len(errors)} within 0.5pt")

        if within > 20:  # Show detail for promising results
            print(f"  {'Font':<12} {'Size':>5} {'COM':>7} {'Pred':>7} {'err':>7}")
            for fid, size, com_gap, predicted, err in details:
                marker = " !" if abs(err) > 1.0 else ""
                print(f"    {fid:<12} {size:>5} {com_gap:>7.2f} {predicted:>7.2f} {err:>+7.2f}{marker}")


    # ====== CRITICAL TEST: default font at RUN's size, not at 11pt ======
    print("\n\n" + "=" * 100)
    print("HYPOTHESIS: max(run_font, default_font @ RUN_SIZE) -- default scales with run")
    print("=" * 100)

    for dpi in [72, 96, 144, 150]:
        errors = []
        details = []

        for c in com:
            fid = c["font_id"]
            size = c["size"]
            com_gap = c["gap"]
            fm_key = FM_MAP.get(fid)
            if not fm_key or fm_key not in fm:
                continue
            run_win = word_line_height(fm[fm_key], size, dpi)
            run_hhea = word_hhea_line_height(fm[fm_key], size, dpi)
            run_h = max(run_win, run_hhea)

            def_win = word_line_height(cambria, size, dpi)
            def_hhea = word_hhea_line_height(cambria, size, dpi)
            def_h = max(def_win, def_hhea)

            predicted = max(run_h, def_h)
            err = predicted - com_gap
            errors.append(abs(err))
            details.append((fid, size, com_gap, predicted, err))

        avg_err = sum(errors) / len(errors)
        max_err = max(errors)
        within = sum(1 for e in errors if e <= 0.5)
        print(f"  DPI={dpi:>4}: avg |err|={avg_err:.3f}pt, max={max_err:.3f}pt, {within}/{len(errors)} within 0.5pt")

        if dpi in (72, 150):
            print(f"  {'Font':<12} {'Size':>5} {'COM':>7} {'Pred':>7} {'err':>7}")
            for fid, size, com_gap, predicted, err in details:
                marker = " !" if abs(err) > 1.0 else ""
                print(f"    {fid:<12} {size:>5} {com_gap:>7.2f} {predicted:>7.2f} {err:>+7.2f}{marker}")

    # ====== BEST CANDIDATE: max(run, max(default_11, default_runsize)) ======
    print("\n\n" + "=" * 100)
    print("HYPOTHESIS: max(run, default@11pt, default@run_size) -- both defaults")
    print("=" * 100)

    for dpi in [72, 96, 144, 150]:
        errors = []
        details = []
        def11 = word_max_line_height(cambria, 11.0, dpi)

        for c in com:
            fid = c["font_id"]
            size = c["size"]
            com_gap = c["gap"]
            fm_key = FM_MAP.get(fid)
            if not fm_key or fm_key not in fm:
                continue
            run_h = word_max_line_height(fm[fm_key], size, dpi)
            def_at_sz = word_max_line_height(cambria, size, dpi)
            predicted = max(run_h, def11, def_at_sz)
            err = predicted - com_gap
            errors.append(abs(err))
            details.append((fid, size, com_gap, run_h, def11, def_at_sz, predicted, err))

        avg_err = sum(errors) / len(errors)
        max_err = max(errors)
        within = sum(1 for e in errors if e <= 0.5)
        print(f"  DPI={dpi:>4}: avg |err|={avg_err:.3f}pt, max={max_err:.3f}pt, {within}/{len(errors)} within 0.5pt")

        if dpi == 72:
            print(f"  {'Font':<12} {'Size':>5} {'COM':>7} {'run':>7} {'d11':>7} {'d@sz':>7} {'Pred':>7} {'err':>7}")
            for fid, size, com_gap, run_h, d11, d_sz, predicted, err in details:
                marker = " !" if abs(err) > 1.0 else ""
                print(f"    {fid:<12} {size:>5} {com_gap:>7.2f} {run_h:>7.2f} {d11:>7.2f} {d_sz:>7.2f} {predicted:>7.2f} {err:>+7.2f}{marker}")

    # ====== Parameter fitting: what default height per size matches COM? ======
    print("\n\n" + "=" * 100)
    print("PARAMETER FITTING: what 'default' value per font-size matches COM?")
    print("(for fonts where run_height < COM, the default must be >= COM)")
    print("=" * 100)

    for size in [10.5, 11, 12, 14]:
        western_com = None
        ea_com = {}
        for c in com:
            if c["size"] != size:
                continue
            fm_key = FM_MAP.get(c["font_id"])
            if not fm_key or fm_key not in fm:
                continue
            ratio = oxi_line_height_ratio(fm[fm_key])
            run_h = ratio * size
            if c["font_id"] in {"arial", "calibri", "century", "tnr"}:
                if western_com is None:
                    western_com = c["gap"]
            ea_com[c["font_id"]] = c["gap"]

        if western_com:
            print(f"  Size {size:>5}pt: Western COM = {western_com:.2f}pt  (ratio = {western_com/size:.4f})")
            # What font-metric ratio would produce this?
            for fname in ["Cambria", "Calibri", "Times New Roman"]:
                if fname in fm:
                    f = fm[fname]
                    win_r = (f["win_ascent"] + f["win_descent"]) / f["units_per_em"]
                    hhea_r = (f.get("ascender", 0) + abs(f.get("descender", 0)) + max(0, f.get("line_gap", 0))) / f["units_per_em"]
                    print(f"    {fname}: win*sz={win_r*size:.3f}, hhea*sz={hhea_r*size:.3f}")
            for ea_fid, ea_gap in ea_com.items():
                if ea_fid in ea_fonts:
                    print(f"    {ea_fid}: COM={ea_gap:.2f} (ratio={ea_gap/size:.4f})")


if __name__ == "__main__":
    main()
