"""Verify: does Word's default font minimum use run's font size or doc default size?

Test approach: Create documents where run font != default font, at different sizes.
Measure line height via COM. Compare against:
  A) max(gdi(run_font, run_size), gdi(default_font, run_size))    -- default@run_size
  B) max(gdi(run_font, run_size), gdi(default_font, default_size)) -- default@default_size

Use existing golden-test documents with known font/size combinations."""
import win32com.client
import os
import time
import math

# We'll use a test document approach: create temp docs with specific font combos
word = win32com.client.Dispatch("Word.Application")
word.Visible = False

def gdi_line_height_96(win_asc_norm, win_des_norm, hhea_asc_norm, hhea_des_norm, hhea_gap_norm, font_size):
    """Replicate Word's line height formula (derived from COM observation) at DPI=96."""
    # zero-line-gap adjustment check (simplified)
    win_total = win_asc_norm + win_des_norm
    if hhea_gap_norm < 0.001 and abs(win_total - 1.0) < 0.01:
        return font_size * (1.0 + 76.0/256.0)

    ppem = round(font_size * 96.0 / 72.0)
    tm_ascent = round(win_asc_norm * ppem)
    tm_descent = round(win_des_norm * ppem)
    tm_height = tm_ascent + tm_descent

    hhea_total = hhea_asc_norm + hhea_des_norm + hhea_gap_norm
    hhea_excess = max(0.0, hhea_total - win_total)
    tm_ext_leading = round(hhea_excess * ppem)

    total_px = tm_height + tm_ext_leading
    return total_px * 15.0 / 20.0

# Font metrics (normalized to 1em, from font_metrics_compact.json)
# Format: (win_asc, win_des, hhea_asc, hhea_des, hhea_gap)
FONT_METRICS = {
    "Calibri": (0.75244140625, 0.25, 0.7529296875, 0.25, 0.0),
    "Times New Roman": (0.9169921875, 0.2158203125, 0.89404296875, 0.21630859375, 0.0),
    "Arial": (0.9052734375, 0.2119140625, 0.87109375, 0.2197265625, 0.0),
    "Yu Gothic Regular": (1.12353515625, 0.31494140625, 0.87646484375, 0.12353515625, 0.0),
    "MS Gothic": (0.859375, 0.140625, 0.859375, 0.140625, 0.0),  # winTotal=1.0, gap=0 -> zero-gap fix
    "MS Mincho": (0.859375, 0.140625, 0.859375, 0.140625, 0.0),  # winTotal=1.0, gap=0 -> zero-gap fix
}

def gdi_h(font_name, font_size):
    m = FONT_METRICS[font_name]
    return gdi_line_height_96(m[0], m[1], m[2], m[3], m[4], font_size)

try:
    # Create test documents with different font/size combos
    # Default font = Calibri at various default sizes
    test_cases = []

    for default_size in [10.5, 11.0]:
        for run_font in ["Yu Gothic Regular", "Arial", "Times New Roman", "MS Gothic"]:
            for run_size in [10.5, 12.0, 14.0, 20.0]:
                test_cases.append((default_size, "Calibri", run_font, run_size))

    print(f"{'DefSize':>7} {'DefFont':<10} {'RunFont':<20} {'RunSize':>7} {'COM_LH':>8} {'A(@run)':>8} {'B(@def)':>8} {'errA':>7} {'errB':>7} {'winner':>7}")
    print("-" * 110)

    total_err_a = 0
    total_err_b = 0
    count = 0

    for (default_size, default_font, run_font, run_size) in test_cases:
        # Create a new document
        doc = word.Documents.Add()
        time.sleep(0.3)

        # Set default font
        # Use wdStyleNormal = -1
        normal_style = doc.Styles(-1)
        normal_style.Font.Name = default_font
        normal_style.Font.Size = default_size

        # Set document grid to "lines" type with a pitch
        sec = doc.Sections(1)
        sec.PageSetup.LayoutMode = 1  # wdLayoutModeLineGrid

        # Disable grid snap for pure measurement
        # Actually we want snap=false to isolate base height
        para = doc.Paragraphs(1)
        para.Format.DisableLineHeightGrid = True  # snap_to_grid = false

        # Set the run font
        rng = para.Range
        rng.Text = "ABCDあいうえお漢字テスト"
        rng.Font.Name = run_font

        # Map "Yu Gothic Regular" to the actual Windows font name
        font_map = {
            "Yu Gothic Regular": "游ゴシック",
            "MS Gothic": "ＭＳ ゴシック",
            "MS Mincho": "ＭＳ 明朝",
        }
        actual_font = font_map.get(run_font, run_font)
        rng.Font.Name = actual_font
        rng.Font.Size = run_size

        time.sleep(0.2)

        # Measure line height via COM
        # Place cursor at start, measure Y, then move down one line
        word.Selection.SetRange(rng.Start, rng.Start)
        y1 = float(word.Selection.Information(6))  # wdVerticalPositionRelativeToPage

        # Add a second paragraph to measure delta
        rng2 = doc.Range(rng.End, rng.End)
        rng2.InsertAfter("\rSecondLine")
        rng2 = doc.Paragraphs(2).Range
        rng2.Font.Name = actual_font
        rng2.Font.Size = run_size
        doc.Paragraphs(2).Format.DisableLineHeightGrid = True

        word.Selection.SetRange(rng2.Start, rng2.Start)
        y2 = float(word.Selection.Information(6))

        com_lh = y2 - y1

        # Calculate predictions
        run_h = gdi_h(run_font, run_size)
        def_h_at_run = gdi_h(default_font, run_size)
        def_h_at_def = gdi_h(default_font, default_size)

        pred_a = max(run_h, def_h_at_run)   # default @ run size
        pred_b = max(run_h, def_h_at_def)   # default @ default size

        err_a = abs(pred_a - com_lh)
        err_b = abs(pred_b - com_lh)

        winner = "A(@run)" if err_a < err_b else "B(@def)" if err_b < err_a else "TIE"

        print(f"{default_size:>7.1f} {default_font:<10} {run_font:<20} {run_size:>7.1f} {com_lh:>8.2f} {pred_a:>8.2f} {pred_b:>8.2f} {err_a:>7.2f} {err_b:>7.2f} {winner:>7}")

        total_err_a += err_a
        total_err_b += err_b
        count += 1

        doc.Close(False)

    print(f"\n{'TOTAL':>56} {total_err_a:>7.2f} {total_err_b:>7.2f}")
    print(f"{'AVG':>56} {total_err_a/count:>7.3f} {total_err_b/count:>7.3f}")
    print(f"\nWinner: {'A (default@run_size)' if total_err_a < total_err_b else 'B (default@default_size)'}")

finally:
    word.Quit()
    print("\nDone.")
