"""Precise Yu Gothic line height measurement: 30 paragraphs per size."""
import win32com.client
import os
import time
import subprocess

subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

def muldiv(a, b, c):
    return (a * b + c // 2) // c

# Yu Gothic: UPM=2048, winA=2017, winD=619, hheaA=1802, hheaD=455, hheaG=1024
UPM, winA, winD = 2048, 2017, 619
hheaA, hheaD, hheaG = 1802, 455, 1024
excess = (hheaA + hheaD + hheaG) - (winA + winD)  # 645

SIZES = [9.0, 9.5, 10.0, 10.5, 11.0, 11.5, 12.0, 13.0, 14.0, 16.0, 18.0, 20.0, 24.0]
N_PARAS = 30

# Phase 1: Create documents (one per size to avoid page overflow)
test_dir = os.path.abspath("_yg_test")
os.makedirs(test_dir, exist_ok=True)

print("Phase 1: Creating test documents...")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    for sz in SIZES:
        doc = word.Documents.Add()
        time.sleep(0.3)

        ns = doc.Styles(-1)
        ns.Font.Name = "Calibri"
        ns.Font.Size = 10.5
        ns.ParagraphFormat.SpaceBefore = 0
        ns.ParagraphFormat.SpaceAfter = 0

        # Build paragraphs
        first = True
        for i in range(N_PARAS):
            if not first:
                doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
            first = False
            pn = doc.Paragraphs.Count
            doc.Paragraphs(pn).Range.Text = f"YG {sz}pt L{i+1}"

        time.sleep(0.3)

        for pi in range(1, N_PARAS + 1):
            p = doc.Paragraphs(pi)
            p.Format.DisableLineHeightGrid = True
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Format.LineSpacingRule = 0
            p.Range.Font.Name = "\u6E38\u30B4\u30B7\u30C3\u30AF"
            p.Range.Font.Size = sz

        fpath = os.path.join(test_dir, f"yg_{sz}.docx")
        doc.SaveAs2(fpath)
        doc.Close(False)
        print(f"  {sz}pt: {N_PARAS} paragraphs")
finally:
    word.Quit()
    time.sleep(2)

# Phase 2: Measure
print("\nPhase 2: Measuring...")
subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

print(f"\n{'Size':>5} {'ppem':>4} {'COM_avg':>8} {'COM_med':>8} {'GDI@96':>8} {'diff':>7} {'diff_px':>8} {'twips':>6}")
print("-" * 65)

try:
    for sz in SIZES:
        fpath = os.path.join(test_dir, f"yg_{sz}.docx")
        doc = word.Documents.Open(fpath)
        time.sleep(0.5)

        ys = []
        for pi in range(1, N_PARAS + 1):
            word.Selection.SetRange(doc.Paragraphs(pi).Range.Start,
                                    doc.Paragraphs(pi).Range.Start)
            y = float(word.Selection.Information(6))
            ys.append(y)

        # Deltas
        deltas = [ys[i+1] - ys[i] for i in range(len(ys)-1)]
        avg = sum(deltas) / len(deltas)
        deltas_sorted = sorted(deltas)
        med = deltas_sorted[len(deltas_sorted) // 2]

        # GDI formula
        ppem = round(sz * 96.0 / 72.0)
        if hheaG == 0 and (winA + winD) == UPM:
            gdi = sz * 1.296875
        else:
            tmH = muldiv(winA, ppem, UPM) + muldiv(winD, ppem, UPM)
            tmEL = muldiv(excess, ppem, UPM)
            gdi = (tmH + tmEL) * 15.0 / 20.0

        diff = avg - gdi
        diff_px = diff * 96.0 / 72.0
        twips = round(avg * 20)

        print(f"{sz:>5.1f} {ppem:>4} {avg:>8.3f} {med:>8.2f} {gdi:>8.2f} {diff:>+7.3f} {diff_px:>+8.3f} {twips:>6}")

        doc.Close(False)

finally:
    word.Quit()

# Cleanup
import shutil
shutil.rmtree(test_dir, ignore_errors=True)
print("\nDone.")
