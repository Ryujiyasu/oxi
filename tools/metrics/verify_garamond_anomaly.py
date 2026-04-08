"""Round 27: Garamond 10.5pt -0.5pt anomaly investigation.

Round 26 found Garamond 10.5pt: raw=(18-12)/2=3.0, measured=2.5, delta=-0.5.
All other LM2 first-para residuals are ±0.25. This -0.5 is anomalous.

Hypotheses:
  H1: Garamond LM0_lh at 10.5pt is actually 13.0pt (not 12.0), measurement noise upstream
  H2: Word uses a different "centering height" for Garamond (e.g., emHeight or sCapHeight)
  H3: Garamond at 10.5pt triggers a different code path (e.g., min font size for LM2)
  H4: Sub-pt pixel-snap quantization that happens to land at 0.5pt only here

Test approach:
  1. Re-measure Garamond at 10.5pt LM0 5 times to confirm 12.0pt
  2. Sweep Garamond sub-pt sizes: 10, 10.25, 10.5, 10.75, 11
  3. Compare with Garamond Bold and other Adobe Garamond variants if installed
  4. Check Garamond ascent/descent vs other Latin fonts via DocumentClass.Font properties
"""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure(font, size, lm):
    doc = word.Documents.Add()
    time.sleep(0.15)
    ps = doc.PageSetup
    ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
    ps.TopMargin = 72; ps.BottomMargin = 72
    try: ps.LayoutMode = lm
    except: pass
    rng = doc.Range()
    rng.InsertAfter("ABC\nDEF")
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font
    time.sleep(0.1)
    try:
        y1 = doc.Paragraphs(1).Range.Information(6)
        y2 = doc.Paragraphs(2).Range.Information(6)
        # Also probe for actual font name (substitution check)
        actual = doc.Paragraphs(1).Range.Characters(1).Font.Name
    except Exception:
        doc.Close(SaveChanges=False); return None
    doc.Close(SaveChanges=False)
    return (y1, y2 - y1, actual)

# Test 1: Repeatability of Garamond 10.5pt LM0
print("=== Test 1: Garamond 10.5pt LM0 repeatability (5 trials) ===")
for trial in range(5):
    r = measure("Garamond", 10.5, 0)
    if r: print(f"  trial{trial+1}: y1={r[0]} h={r[1]} actual_font={r[2]!r}")

# Test 2: sub-pt sweep
print("\n=== Test 2: Garamond sub-pt sweep LM0 + LM2 ===")
print(f"{'size':<7} {'LM0_y':<7} {'LM0_h':<7} {'LM2_y':<7} {'LM2_h':<7} {'offset':<7} {'raw':<7} {'delta':<6} {'actual':<25}")
for size in [9, 9.5, 10, 10.25, 10.5, 10.75, 11, 11.5, 12]:
    r0 = measure("Garamond", size, 0)
    r2 = measure("Garamond", size, 2)
    if r0 is None or r2 is None: continue
    lm0_y, lm0_h, _ = r0
    lm2_y, lm2_h, actual = r2
    offset = lm2_y - 72
    raw = (lm2_h - lm0_h) / 2
    delta = offset - raw
    print(f"{size:<7} {lm0_y:<7.2f} {lm0_h:<7.2f} {lm2_y:<7.2f} {lm2_h:<7.2f} {offset:<7.2f} {raw:<7.3f} {delta:<6.2f} {actual!r:<25}")

# Test 3: Other 10.5pt Latin fonts at LM0 to see if 12.0 is unique
print("\n=== Test 3: 10.5pt LM0_h across Latin fonts ===")
for font in ["Garamond", "Times New Roman", "Times", "Calibri", "Arial", "Cambria", "Constantia", "Georgia", "Verdana", "Century"]:
    r = measure(font, 10.5, 0)
    if r:
        print(f"  {font:<22} LM0_h={r[1]:.2f} actual={r[2]!r}")

word.Quit()
