"""Measure line heights for multiple CJK fonts to find K values."""
import win32com.client
import os
import time
import subprocess
import shutil

subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

def muldiv(a, b, c):
    return (a * b + c // 2) // c

FONTS = [
    ("\u6E38\u30B4\u30B7\u30C3\u30AF", "Yu Gothic"),      # Yu Gothic
    ("\u6E38\u660E\u671D", "Yu Mincho"),                    # Yu Mincho
    ("\uFF2D\uFF33 \u30B4\u30B7\u30C3\u30AF", "MS Gothic"),# MS Gothic
    ("\uFF2D\uFF33 \u660E\u671D", "MS Mincho"),             # MS Mincho
    ("Meiryo", "Meiryo"),
]

SIZES = [9.0, 9.5, 10.0, 10.5, 11.0, 11.5, 12.0, 13.0, 14.0]
N_PARAS = 15  # fewer to avoid page overflow

test_dir = os.path.abspath("_multi_cjk_test")
os.makedirs(test_dir, exist_ok=True)

# Phase 1: Create one docx per font per size
print("Phase 1: Creating test documents...")
for font_name, font_label in FONTS:
    subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
    time.sleep(2)

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

            # Disable grid
            doc.Sections(1).PageSetup.LayoutMode = 0

            first = True
            for i in range(N_PARAS):
                if not first:
                    doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
                first = False
                pn = doc.Paragraphs.Count
                doc.Paragraphs(pn).Range.Text = f"{font_label} {sz}pt L{i+1}"

            time.sleep(0.2)

            for pi in range(1, N_PARAS + 1):
                p = doc.Paragraphs(pi)
                p.Format.DisableLineHeightGrid = True
                p.Format.SpaceBefore = 0
                p.Format.SpaceAfter = 0
                p.Format.LineSpacingRule = 0  # Single
                p.Range.Font.Name = font_name
                p.Range.Font.Size = sz

            fpath = os.path.join(test_dir, f"{font_label}_{sz}.docx")
            doc.SaveAs2(fpath)
            doc.Close(False)
        print(f"  {font_label}: done")
    except Exception as e:
        print(f"  {font_label}: ERROR {e}")
    finally:
        word.Quit()

# Phase 2: Measure
print("\nPhase 2: Measuring...")
subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

results = {}  # font_label -> [(size, twips)]

for font_name, font_label in FONTS:
    subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
    time.sleep(2)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    font_results = []
    try:
        for sz in SIZES:
            fpath = os.path.join(test_dir, f"{font_label}_{sz}.docx")
            if not os.path.exists(fpath):
                font_results.append((sz, None))
                continue

            doc = word.Documents.Open(fpath)
            time.sleep(0.3)

            ys = []
            for pi in range(1, N_PARAS + 1):
                word.Selection.SetRange(doc.Paragraphs(pi).Range.Start,
                                        doc.Paragraphs(pi).Range.Start)
                y = float(word.Selection.Information(6))
                ys.append(y)

            deltas = [ys[i+1] - ys[i] for i in range(len(ys)-1)]
            # Filter out page breaks (negative deltas or huge jumps)
            good = [d for d in deltas if 5 < d < 40]
            if good:
                avg = sum(good) / len(good)
                twips = round(avg * 20)
            else:
                twips = None
            font_results.append((sz, twips))
            doc.Close(False)

        results[font_label] = font_results
        print(f"  {font_label}: measured")
    except Exception as e:
        print(f"  {font_label}: ERROR {e}")
        results[font_label] = [(sz, None) for sz in SIZES]
    finally:
        word.Quit()

# Phase 3: Analysis
print("\n" + "=" * 80)
print("RESULTS")
print("=" * 80)

# Font metrics (from JSON)
METRICS = {
    "Yu Gothic": {"UPM": 2048, "winA": 2017, "winD": 619, "hA": 1802, "hD": 455, "hG": 1024, "tA": 1802, "tD": 246, "tG": 1024},
    "Yu Mincho": {"UPM": 2048, "winA": 2038, "winD": 598, "hA": 1802, "hD": 455, "hG": 1024, "tA": 1802, "tD": 246, "tG": 1024},
    "MS Gothic": {"UPM": 256, "winA": 220, "winD": 36, "hA": 220, "hD": 36, "hG": 0, "tA": 220, "tD": 36, "tG": 0},
    "MS Mincho": {"UPM": 256, "winA": 220, "winD": 36, "hA": 220, "hD": 36, "hG": 0, "tA": 220, "tD": 36, "tG": 0},
    "Meiryo":    {"UPM": 2048, "winA": 2146, "winD": 555, "hA": 2146, "hD": 555, "hG": 0, "tA": 1946, "tD": 455, "tG": 0},  # approximate
}

for font_label in [fl for _, fl in FONTS]:
    data = results.get(font_label, [])
    if not data:
        continue

    m = METRICS.get(font_label)
    if not m:
        print(f"\n{font_label}: no metrics available")
        continue

    UPM = m["UPM"]
    excess = max(0, (m["hA"] + m["hD"] + m["hG"]) - (m["winA"] + m["winD"]))
    gdi_K = m["winA"] + m["winD"] + excess

    print(f"\n{font_label} (UPM={UPM}, winA={m['winA']}, winD={m['winD']}, excess={excess}, GDI_K={gdi_K})")
    print(f"  typoA={m['tA']} typoD={m['tD']} typoG={m['tG']}")
    print(f"  {'Size':>5} {'COM_tw':>7} {'GDI_tw':>7} {'diff':>5} {'K_eff':>8}")

    k_values = []
    for sz, tw in data:
        if tw is None:
            print(f"  {sz:>5.1f}    N/A")
            continue
        fs_tw = round(sz * 20)
        gdi_tw = muldiv(gdi_K, fs_tw, UPM)
        diff = tw - gdi_tw
        k_eff = (tw * UPM - UPM // 2) / fs_tw
        k_values.append(k_eff)
        print(f"  {sz:>5.1f} {tw:>7} {gdi_tw:>7} {diff:>+5} {k_eff:>8.1f}")

    if k_values:
        avg_k = sum(k_values) / len(k_values)
        print(f"  Average K_eff = {avg_k:.1f}")
        # Find best integer K
        best_K = round(avg_k)
        total_err = 0
        for sz, tw in data:
            if tw is None: continue
            fs_tw = round(sz * 20)
            pred = muldiv(best_K, fs_tw, UPM)
            total_err += abs(pred - tw)
        print(f"  Best integer K = {best_K}, total_err = {total_err}")

        # Check nearby
        for k_try in range(best_K - 5, best_K + 6):
            err = 0
            for sz, tw in data:
                if tw is None: continue
                fs_tw = round(sz * 20)
                pred = muldiv(k_try, fs_tw, UPM)
                err += abs(pred - tw)
            if err <= 2:
                print(f"    K={k_try}: err={err}")

# Cleanup
shutil.rmtree(test_dir, ignore_errors=True)
print("\nDone.")
