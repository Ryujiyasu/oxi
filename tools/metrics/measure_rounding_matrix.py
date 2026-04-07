"""Comprehensive rounding measurement: multiple fonts × sizes × factors at LM0"""
import win32com.client
import time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    time.sleep(1)

    try:
        fonts_sizes = [
            ("Calibri", 11), ("Calibri", 14), ("Calibri", 26),
            ("Cambria", 11), ("Cambria", 14),
        ]
        factors_line = [
            (1.0, 240),    # Single
            (1.15, 276),   # 276/240
            (1.5, 360),    # 360/240
            (2.0, 480),    # 480/240
        ]

        print(f"{'Font':>10s} {'Size':>4s} {'Factor':>6s} {'Gap':>7s} {'GapTw':>6s} | "
              f"{'pre+ceil':>8s} {'raw+flr':>8s} {'Match':>8s}")
        print("-" * 80)

        for font, size in fonts_sizes:
            # Get metrics
            upm_map = {"Calibri": 2048, "Cambria": 2048}
            wa_map = {"Calibri": 1950, "Cambria": 1946}
            wd_map = {"Calibri": 550, "Cambria": 455}
            upm = upm_map[font]
            win_sum = (wa_map[font] + wd_map[font]) / upm
            raw_base = win_sum * size
            pre_base_tw = int(raw_base * 20 / 10) * 10  # floor to 10tw
            pre_base = pre_base_tw / 20.0

            for factor, line_val in factors_line:
                doc = word.Documents.Add()
                time.sleep(0.2)
                sec = doc.Sections(1)
                sec.PageSetup.LayoutMode = 0

                rng = doc.Range(0, 0)
                rng.Text = "A\nB\n"
                rng.Font.Name = font
                rng.Font.Size = size

                for i in range(1, doc.Paragraphs.Count + 1):
                    p = doc.Paragraphs(i)
                    if factor == 1.0:
                        p.Format.LineSpacingRule = 0  # Single
                    else:
                        p.Format.LineSpacingRule = 5  # Multiple
                        p.Format.LineSpacing = factor * 12
                    p.Format.SpaceBefore = 0
                    p.Format.SpaceAfter = 0

                doc.Repaginate()
                time.sleep(0.2)

                y1 = doc.Paragraphs(1).Range.Information(6)
                y2 = doc.Paragraphs(2).Range.Information(6)
                gap = y2 - y1
                gap_tw = gap * 20

                # Pre-rounded + ceil
                spaced_pre = pre_base * factor
                pre_ceil_tw = int((spaced_pre * 20 + 9.999) // 10) * 10
                pre_ceil = pre_ceil_tw / 20.0

                # Raw + floor
                spaced_raw = raw_base * factor
                raw_floor_tw = int(spaced_raw * 20 / 10) * 10
                raw_floor = raw_floor_tw / 20.0

                match = ""
                if abs(gap_tw - pre_ceil_tw) < 1: match = "pre+ceil"
                elif abs(gap_tw - raw_floor_tw) < 1: match = "raw+flr"
                else: match = "NONE"

                print(f"{font:>10s} {size:>4d} {factor:>6.2f} {gap:>7.2f} {gap_tw:>6.0f} | "
                      f"{pre_ceil:>8.2f} {raw_floor:>8.2f} {match:>8s}")

                doc.Close(0)

    finally:
        word.Quit()

if __name__ == "__main__":
    measure()
