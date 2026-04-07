"""Verify floor vs ceil rounding for LM≥1 Multiple spacing with various factors"""
import win32com.client
import time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    time.sleep(1)

    try:
        # Test multiple factors at LM0 and LM1
        factors = [1.15, 1.5, 2.0, 1.08, 1.3]

        for lm in [0, 1]:
            print(f"\n{'='*50}")
            print(f"LayoutMode={lm}")
            print(f"{'='*50}")

            for factor in factors:
                doc = word.Documents.Add()
                time.sleep(0.3)
                sec = doc.Sections(1)
                sec.PageSetup.LayoutMode = lm

                rng = doc.Range(0, 0)
                rng.Text = "Line1\nLine2\nLine3\n"
                rng.Font.Name = "Calibri"
                rng.Font.Size = 11

                line_val = factor * 12  # 12pt = standard Single line
                for i in range(1, doc.Paragraphs.Count + 1):
                    p = doc.Paragraphs(i)
                    p.Format.LineSpacingRule = 5  # Multiple
                    p.Format.LineSpacing = line_val
                    p.Format.SpaceBefore = 0
                    p.Format.SpaceAfter = 0

                doc.Repaginate()
                time.sleep(0.2)

                y1 = doc.Paragraphs(1).Range.Information(6)
                y2 = doc.Paragraphs(2).Range.Information(6)
                gap = y2 - y1
                gap_tw = gap * 20

                # Compute expected ceil and floor
                if lm == 0:
                    base = 13.0  # Calibri 11pt LM0
                else:
                    base = 18.0  # Calibri 11pt LM1 grid-snapped Single

                raw = base * factor
                raw_tw = raw * 20
                ceil_tw = int((raw_tw + 9.999) // 10) * 10  # ceil to 10
                floor_tw = int(raw_tw // 10) * 10  # floor to 10

                match = ""
                if abs(gap_tw - ceil_tw) < 1: match = "CEIL"
                elif abs(gap_tw - floor_tw) < 1: match = "FLOOR"
                else: match = "???"

                print(f"  {factor:.2f}x: gap={gap:.2f}pt ({gap_tw:.0f}tw) | base*f={raw:.4f} → ceil={ceil_tw/20:.2f} floor={floor_tw/20:.2f} | {match}")

                doc.Close(0)

    finally:
        word.Quit()

if __name__ == "__main__":
    measure()
