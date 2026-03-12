"""Two-phase: create docx files first, then measure all at once."""
import win32com.client
import ctypes
import os
import time
import subprocess
import json

subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

# System DPI
try:
    user32 = ctypes.windll.user32
    user32.SetProcessDPIAware()
    hdc = user32.GetDC(0)
    gdi32 = ctypes.windll.gdi32
    dpi_y = gdi32.GetDeviceCaps(hdc, 90)
    user32.ReleaseDC(0, hdc)
    print(f"System DPI: {dpi_y} ({dpi_y/96*100:.0f}%)")
except:
    dpi_y = 96

def muldiv_int(a, b, c):
    return (a * b + c // 2) // c

FONTS = [
    ("Calibri", 2048, 1950, 550, 1536, 512, 452, "Calibri"),
    ("TNR", 2048, 1825, 443, 1825, 443, 87, "Times New Roman"),
    ("Arial", 2048, 1854, 434, 1854, 434, 67, "Arial"),
    ("YuGothic", 2048, 2017, 619, 1802, 455, 1024, "\u6E38\u30B4\u30B7\u30C3\u30AF"),
    ("MSGothic", 256, 220, 36, 220, 36, 0, "\uFF2D\uFF33 \u30B4\u30B7\u30C3\u30AF"),
    ("YuMincho", 2048, 2017, 619, 1802, 455, 1024, "\u6E38\u660E\u671D"),
]
SIZES = [9.0, 10.0, 10.5, 11.0, 12.0, 14.0, 16.0, 20.0, 24.0]

def gdi_lh(upm, winA, winD, hheaA, hheaD, hheaG, sz):
    if hheaG == 0 and (winA + winD) == upm:
        return sz * (1.0 + 76.0/256.0)
    ppem = round(sz * 96.0 / 72.0)
    tmH = muldiv_int(winA, ppem, upm) + muldiv_int(winD, ppem, upm)
    excess = max(0, (hheaA + hheaD + hheaG) - (winA + winD))
    tmExt = muldiv_int(excess, ppem, upm)
    return (tmH + tmExt) * 15.0 / 20.0

# Phase 1: Create one docx per font
test_dir = os.path.abspath("_lh_test")
os.makedirs(test_dir, exist_ok=True)

print("\nPhase 1: Creating test documents...")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    for (key, upm, winA, winD, hheaA, hheaD, hheaG, wname) in FONTS:
        doc = word.Documents.Add()
        time.sleep(0.5)

        # Set Normal style
        ns = doc.Styles(-1)
        ns.Font.Name = "Calibri"
        ns.Font.Size = 10.5
        ns.ParagraphFormat.SpaceBefore = 0
        ns.ParagraphFormat.SpaceAfter = 0

        # Build paragraphs: 3 per size (for averaging)
        first = True
        for sz in SIZES:
            for line_idx in range(3):
                if not first:
                    doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
                first = False
                para_num = doc.Paragraphs.Count
                p = doc.Paragraphs(para_num)
                p.Range.Text = f"Test {key} {sz}pt L{line_idx+1}"

        time.sleep(0.5)

        # Format all paragraphs
        pi = 1
        for sz in SIZES:
            for line_idx in range(3):
                p = doc.Paragraphs(pi)
                p.Format.DisableLineHeightGrid = True
                p.Format.SpaceBefore = 0
                p.Format.SpaceAfter = 0
                p.Format.LineSpacingRule = 0
                p.Range.Font.Name = wname
                p.Range.Font.Size = sz
                pi += 1

        fpath = os.path.join(test_dir, f"test_{key}.docx")
        doc.SaveAs2(fpath)
        doc.Close(False)
        print(f"  {key}: {pi-1} paragraphs -> {fpath}")
finally:
    word.Quit()
    time.sleep(2)

# Phase 2: Measure
print("\nPhase 2: Measuring...")
subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

print(f"\n{'Font':<13} {'Size':>5} {'COM_d12':>8} {'COM_d23':>8} {'COM_avg':>8} {'GDI@96':>8} {'err':>7}")
print("-" * 70)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    for (key, upm, winA, winD, hheaA, hheaD, hheaG, wname) in FONTS:
        fpath = os.path.join(test_dir, f"test_{key}.docx")
        if not os.path.exists(fpath):
            print(f"{key:<13} SKIP (file not found)")
            continue

        doc = word.Documents.Open(fpath)
        time.sleep(0.5)

        total_paras = doc.Paragraphs.Count
        pi = 1
        for sz in SIZES:
            ys = []
            for line_idx in range(3):
                if pi > total_paras:
                    break
                word.Selection.SetRange(doc.Paragraphs(pi).Range.Start,
                                        doc.Paragraphs(pi).Range.Start)
                y = float(word.Selection.Information(6))
                ys.append(y)
                pi += 1

            if len(ys) == 3:
                d12 = ys[1] - ys[0]
                d23 = ys[2] - ys[1]
                avg = (d12 + d23) / 2.0
                predicted = gdi_lh(upm, winA, winD, hheaA, hheaD, hheaG, sz)
                err = abs(predicted - avg)
                print(f"{key:<13} {sz:>5.1f} {d12:>8.2f} {d23:>8.2f} {avg:>8.2f} {predicted:>8.2f} {err:>7.2f}")
            else:
                print(f"{key:<13} {sz:>5.1f} INCOMPLETE")

        doc.Close(False)
        print()

finally:
    word.Quit()

# Cleanup
import shutil
shutil.rmtree(test_dir, ignore_errors=True)
print("Done.")
