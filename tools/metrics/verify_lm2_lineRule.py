"""Round 28: LM2 first-paragraph behavior with explicit lineRule != auto.

Round 23/24 confirmed the closed-form
  P0_y = topMargin + (P0_h - LM0_lh) / 2  with  P0_h = (floor(LM0_lh/pitch)+1)*pitch
for default lineSpacingRule = auto/single.

Open: how does LM2 first-para Y change for explicit lineRule values:
  - "multiple" (e.g. 1.15x = w:line=276, 1.5x = 360, 2.0x = 480)
  - "atLeast" (e.g. w:line=240 lineRule=atLeast → at least 12pt)
  - "exact"   (e.g. w:line=240 lineRule=exact → exactly 12pt)

Hypothesis A: same formula, but LM0_lh is replaced by the rule-effective line height
Hypothesis B: lineRule disables LM2 grid centering (P0_y = topMargin directly)
Hypothesis C: only the FIRST line of the para is centered, but subsequent lines use rule

Test: vary lineRule × line value × font, measure P0_y, P0_h, P1_y, P2_y deltas.
"""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

WD_LINE_SPACING_SINGLE = 0
WD_LINE_SPACING_1PT5 = 1
WD_LINE_SPACING_DOUBLE = 2
WD_LINE_SPACING_AT_LEAST = 3
WD_LINE_SPACING_EXACTLY = 4
WD_LINE_SPACING_MULTIPLE = 5

def measure(font, size, lm, rule, line_value=None):
    """rule = 'single' | '1.15' | '1.5' | 'double' | 'multiple_X' | 'atLeast_Xpt' | 'exact_Xpt'"""
    doc = word.Documents.Add()
    time.sleep(0.15)
    ps = doc.PageSetup
    ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
    ps.TopMargin = 72; ps.BottomMargin = 72
    try: ps.LayoutMode = lm
    except: pass
    rng = doc.Range()
    rng.InsertAfter("Line1\nLine2\nLine3")
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font
    pf = rng.ParagraphFormat
    if rule == "single":
        pf.LineSpacingRule = WD_LINE_SPACING_SINGLE
    elif rule == "1.15":
        pf.LineSpacingRule = WD_LINE_SPACING_MULTIPLE
        pf.LineSpacing = size * 1.15
    elif rule == "1.5":
        pf.LineSpacingRule = WD_LINE_SPACING_1PT5
    elif rule == "double":
        pf.LineSpacingRule = WD_LINE_SPACING_DOUBLE
    elif rule.startswith("multiple_"):
        mul = float(rule.split("_")[1])
        pf.LineSpacingRule = WD_LINE_SPACING_MULTIPLE
        pf.LineSpacing = size * mul
    elif rule.startswith("atLeast_"):
        v = float(rule.split("_")[1])
        pf.LineSpacingRule = WD_LINE_SPACING_AT_LEAST
        pf.LineSpacing = v
    elif rule.startswith("exact_"):
        v = float(rule.split("_")[1])
        pf.LineSpacingRule = WD_LINE_SPACING_EXACTLY
        pf.LineSpacing = v
    time.sleep(0.1)
    try:
        ys = []
        for i in range(1, 4):
            try: ys.append(doc.Paragraphs(i).Range.Information(6))
            except: ys.append(None)
    except Exception:
        doc.Close(SaveChanges=False); return None
    doc.Close(SaveChanges=False)
    return ys

print(f"{'font':<14} {'sz':<5} {'rule':<14} {'P0_y':<7} {'P1_y':<7} {'P2_y':<7} {'P0_h':<7} {'P1_h':<7}")
for font in ["Times New Roman", "ＭＳ 明朝"]:
    for size in [10.5, 12, 14]:
        for rule in [
            "single",
            "1.15",
            "1.5",
            "double",
            "multiple_1.0",
            "multiple_1.25",
            "multiple_2.0",
            "atLeast_15",
            "atLeast_24",
            "atLeast_36",
            "exact_12",
            "exact_18",
            "exact_24",
            "exact_36",
        ]:
            for lm in [2]:
                ys = measure(font, size, lm, rule)
                if ys is None or ys[0] is None: continue
                p0h = (ys[1]-ys[0]) if ys[1] else 0
                p1h = (ys[2]-ys[1]) if ys[2] and ys[1] else 0
                print(f"{font:<14} {size:<5} {rule:<14} {ys[0]:<7.2f} {(ys[1] or 0):<7.2f} {(ys[2] or 0):<7.2f} {p0h:<7.2f} {p1h:<7.2f}")
        print()

word.Quit()
