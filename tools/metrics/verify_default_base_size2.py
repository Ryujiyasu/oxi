"""Verify: default font minimum uses run_size or default_size?
Simplified: one doc per test, robust COM handling."""
import win32com.client
import os
import time
import subprocess

# Kill any lingering Word processes
subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(2)

def gdi_line_height_96(win_asc, win_des, hhea_asc, hhea_des, hhea_gap, font_size):
    win_total = win_asc + win_des
    if hhea_gap < 0.001 and abs(win_total - 1.0) < 0.01:
        return font_size * (1.0 + 76.0/256.0)
    ppem = round(font_size * 96.0 / 72.0)
    tm_h = round(win_asc * ppem) + round(win_des * ppem)
    excess = max(0.0, (hhea_asc + hhea_des + hhea_gap) - win_total)
    tm_ext = round(excess * ppem)
    return (tm_h + tm_ext) * 15.0 / 20.0

METRICS = {
    "Calibri":           (0.75244140625, 0.25, 0.7529296875, 0.25, 0.0),
    "Yu Gothic Regular": (1.12353515625, 0.31494140625, 0.87646484375, 0.12353515625, 0.0),
    "Arial":             (0.9052734375, 0.2119140625, 0.87109375, 0.2197265625, 0.0),
    "Times New Roman":   (0.9169921875, 0.2158203125, 0.89404296875, 0.21630859375, 0.0),
    "MS Gothic":         (0.859375, 0.140625, 0.859375, 0.140625, 0.0),
}

def gdi_h(font, size):
    m = METRICS[font]
    return gdi_line_height_96(*m, size)

FONT_MAP = {
    "Yu Gothic Regular": "游ゴシック",
    "MS Gothic": "ＭＳ ゴシック",
}

# Test cases: (default_size, run_font, run_size)
# Default font is always Calibri
cases = [
    (10.5, "Yu Gothic Regular", 10.5),
    (10.5, "Yu Gothic Regular", 14.0),
    (10.5, "Yu Gothic Regular", 20.0),
    (10.5, "Arial", 10.5),
    (10.5, "Arial", 14.0),
    (10.5, "Arial", 20.0),
    (10.5, "MS Gothic", 10.5),
    (10.5, "MS Gothic", 14.0),
    (11.0, "Yu Gothic Regular", 10.5),
    (11.0, "Yu Gothic Regular", 14.0),
    (11.0, "Yu Gothic Regular", 20.0),
    (11.0, "Arial", 10.5),
    (11.0, "Arial", 14.0),
    (11.0, "Arial", 20.0),
    (11.0, "MS Gothic", 10.5),
    (11.0, "MS Gothic", 14.0),
    # Key discriminating cases: run_size != default_size, and default font height matters
    (10.5, "Times New Roman", 24.0),
    (11.0, "Times New Roman", 24.0),
    (10.5, "Times New Roman", 10.5),
    (11.0, "Times New Roman", 10.5),
]

print(f"{'DefSz':>5} {'RunFont':<20} {'RunSz':>5} {'COM':>8} {'A@run':>8} {'B@def':>8} {'eA':>6} {'eB':>6} {'win':>7} {'run_h':>7} {'d@r':>7} {'d@d':>7}")
print("-" * 120)

total_a = 0.0
total_b = 0.0
n = 0

for (def_sz, run_font, run_sz) in cases:
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Add()
        time.sleep(0.5)

        # Set Normal style default font
        ns = doc.Styles(-1)  # wdStyleNormal
        ns.Font.Name = "Calibri"
        ns.Font.Size = def_sz

        # Set grid mode
        doc.Sections(1).PageSetup.LayoutMode = 1  # wdLayoutModeLineGrid

        # First paragraph
        p1 = doc.Paragraphs(1)
        p1.Format.DisableLineHeightGrid = True
        r1 = p1.Range
        r1.Text = "ABCあいう漢字"
        actual_font = FONT_MAP.get(run_font, run_font)
        r1.Font.Name = actual_font
        r1.Font.Size = run_sz

        # Second paragraph (same style)
        r1.InsertAfter("\r")
        p2 = doc.Paragraphs(2)
        p2.Format.DisableLineHeightGrid = True
        r2 = p2.Range
        r2.Text = "XYZかきく数字"
        r2.Font.Name = actual_font
        r2.Font.Size = run_sz

        time.sleep(0.3)

        word.Selection.SetRange(p1.Range.Start, p1.Range.Start)
        y1 = float(word.Selection.Information(6))
        word.Selection.SetRange(p2.Range.Start, p2.Range.Start)
        y2 = float(word.Selection.Information(6))

        com_lh = y2 - y1

        rh = gdi_h(run_font, run_sz)
        dr = gdi_h("Calibri", run_sz)
        dd = gdi_h("Calibri", def_sz)

        pred_a = max(rh, dr)
        pred_b = max(rh, dd)

        ea = abs(pred_a - com_lh)
        eb = abs(pred_b - com_lh)

        win = "A@run" if ea < eb - 0.01 else "B@def" if eb < ea - 0.01 else "TIE"
        print(f"{def_sz:>5.1f} {run_font:<20} {run_sz:>5.1f} {com_lh:>8.2f} {pred_a:>8.2f} {pred_b:>8.2f} {ea:>6.2f} {eb:>6.2f} {win:>7}  {rh:>6.2f}  {dr:>6.2f}  {dd:>6.2f}")

        total_a += ea
        total_b += eb
        n += 1

        doc.Close(False)
    except Exception as e:
        print(f"{def_sz:>5.1f} {run_font:<20} {run_sz:>5.1f} ERROR: {e}")
    finally:
        try:
            if word:
                word.Quit()
        except:
            pass
        time.sleep(0.5)

if n > 0:
    print(f"\n  Avg error A (default@run_size):     {total_a/n:.3f}pt")
    print(f"  Avg error B (default@default_size): {total_b/n:.3f}pt")
    print(f"  Winner: {'A (default@run_size)' if total_a < total_b else 'B (default@default_size)'}")
print("\nDone.")
