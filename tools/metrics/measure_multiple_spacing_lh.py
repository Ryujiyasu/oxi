"""Measure line height for Multiple spacing (1.15x) paragraphs via COM.
Goal: determine exact formula for MS Mincho 10.5pt * 1.15."""
import win32com.client
import os

docx_path = os.path.abspath(r"tools\golden-test\documents\docx\0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(docx_path, ReadOnly=True)

    # P1: Multiple 1.15x, MS Mincho 10.5pt, empty para
    p1 = doc.Paragraphs(1)
    p2 = doc.Paragraphs(2)
    y1 = p1.Range.Information(6)
    y2 = p2.Range.Information(6)
    gap = y2 - y1

    print(f"P1→P2 gap: {gap:.2f}pt")
    print(f"P1 ls={p1.Format.LineSpacing:.2f} rule={p1.Format.LineSpacingRule}")
    print(f"P1 font={p1.Range.Font.Name} fs={p1.Range.Font.Size}")
    print()

    # Calculate expected values
    # MS Mincho: win_ascent=0.88, win_descent=0.12 (normalized)
    # At 10.5pt: ascent+descent = 10.5 * (win_a + win_d)
    fs = 10.5

    # Try different formulas
    # 1) floor(fs * 83/64 * 8)/8 * 1.15
    cjk_h = int(fs * 83/64 * 8) / 8
    print(f"CJK 83/64 height: {cjk_h:.4f}pt")
    print(f"  × 1.15 = {cjk_h * 1.15:.4f}pt")
    print(f"  ceil to 0.5pt = {((cjk_h * 1.15 * 2).__ceil__()) / 2:.1f}pt")

    # 2) raw win_sum * fs * 1.15
    # MS Mincho: win_ascent=880, win_descent=120, UPM=1000
    # → win_sum = 1.0
    raw = 1.0 * fs * 1.15
    print(f"\nRaw win_sum*fs*1.15: {raw:.4f}pt")
    print(f"  = {raw*20:.1f}tw")
    print(f"  ceil to 10tw: {((raw*20/10).__ceil__())*10/20:.2f}pt")

    # 3) no_grid height * 1.15
    # no_grid = floor(win_sum * fs * 20 / 10) * 10 / 20
    no_grid = int(1.0 * fs * 20 / 10) * 10 / 20
    print(f"\nNo-grid height: {no_grid:.2f}pt")
    print(f"  × 1.15 = {no_grid * 1.15:.4f}pt")
    print(f"  ceil to 0.5pt = {((no_grid * 1.15 * 2).__ceil__()) / 2:.1f}pt")

    # 4) Check cumulative ceil with raw_tw
    raw_tw = raw * 20
    import math
    for j in range(3):
        cn = math.ceil((j+1) * raw_tw / 10) * 10
        cc = math.ceil(j * raw_tw / 10) * 10
        adv = (cn - cc) / 20
        print(f"\nCumul ceil j={j}→{j+1}: advance={adv:.2f}pt (cn={cn} cc={cc})")

    # What raw_tw gives gap=16.0?
    # advance = ceil(1*raw_tw/10)*10 / 20 = 16.0
    # ceil(raw_tw/10)*10 = 320
    # raw_tw/10 must be in (31, 32] → raw_tw in (310, 320]
    print(f"\nFor gap=16.0: raw_tw must be in (310, 320]")
    print(f"Actual raw_tw from formula: {raw_tw:.1f}")

    # 5) Maybe Word uses a different base for Multiple spacing
    # Try: line_height (snapped) * 1.15
    # snapped = 13.5pt for CJK 83/64
    print(f"\nSnapped(13.5) * 1.15 = {13.5 * 1.15:.4f}pt → {13.5*1.15*20:.1f}tw")
    print(f"  ceil(310.5/10)*10 = {math.ceil(310.5/10)*10} → {math.ceil(310.5/10)*10/20:.1f}pt")

    doc.Close(False)
finally:
    word.Quit()
