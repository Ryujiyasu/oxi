"""Test if Exactly mode works correctly for CJK fonts.
Also check document grid settings."""
import win32com.client
import os
import time
import subprocess

subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

test_path = os.path.abspath("_test_exact_cjk.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Add()
    time.sleep(0.5)

    # Check grid settings
    sec = doc.Sections(1)
    ps = sec.PageSetup
    print(f"Page setup:")
    print(f"  LayoutMode = {ps.LayoutMode}")  # 1=LineGrid, 2=Grid, 3=Default
    try:
        print(f"  LinesPage = {ps.LinesPage}")
    except:
        print(f"  LinesPage = N/A")
    try:
        print(f"  LinePitch = {ps.LinePitch}")
    except:
        print(f"  LinePitch = N/A")
    try:
        print(f"  CharsLine = {ps.CharsLine}")
    except:
        print(f"  CharsLine = N/A")

    # Default style info
    ns = doc.Styles(-1)
    print(f"\nNormal style:")
    print(f"  Font: {ns.Font.Name} {ns.Font.Size}pt")
    print(f"  LineSpacingRule={ns.ParagraphFormat.LineSpacingRule}")
    print(f"  LineSpacing={ns.ParagraphFormat.LineSpacing}")

    # Set Normal to Calibri 10.5
    ns.Font.Name = "Calibri"
    ns.Font.Size = 10.5
    ns.ParagraphFormat.SpaceBefore = 0
    ns.ParagraphFormat.SpaceAfter = 0

    # Disable grid
    ps.LayoutMode = 0  # wdLayoutModeDefault (no grid)
    time.sleep(0.3)

    print(f"\nAfter disabling grid:")
    print(f"  LayoutMode = {ps.LayoutMode}")

    # Test: Exactly at various values, for Yu Gothic
    tests = [
        ("YG_Exact30", "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 3, 30.0),
        ("YG_Exact20", "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 3, 20.0),
        ("YG_Exact16.5", "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 3, 16.5),
        ("YG_Exact10", "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 3, 10.0),
        ("YG_Single", "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 0, None),
        ("Cal_Exact20", "Calibri", 10.5, 3, 20.0),
        ("Cal_Exact12.75", "Calibri", 10.5, 3, 12.75),
        ("Cal_Single", "Calibri", 10.5, 0, None),
    ]

    N = 5
    first = True
    for (label, font, sz, rule, spacing) in tests:
        for i in range(N):
            if not first:
                doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
            first = False
            pn = doc.Paragraphs.Count
            doc.Paragraphs(pn).Range.Text = f"{label} L{i+1}"

    time.sleep(0.3)

    pi = 1
    for (label, font, sz, rule, spacing) in tests:
        for i in range(N):
            p = doc.Paragraphs(pi)
            p.Format.DisableLineHeightGrid = True
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Range.Font.Name = font
            p.Range.Font.Size = sz
            p.Format.LineSpacingRule = rule
            if spacing is not None:
                p.Format.LineSpacing = spacing
            pi += 1

    doc.SaveAs2(test_path)
    doc.Close(False)
    print(f"\nSaved: {test_path}")

finally:
    word.Quit()
    time.sleep(2)

# Phase 2: Measure
subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(test_path)
    time.sleep(0.5)

    # Re-check grid
    ps = doc.Sections(1).PageSetup
    print(f"\nRe-opened LayoutMode = {ps.LayoutMode}")

    print(f"\n{'Label':<18} {'Rule':>6} {'Spacing':>8} {'GridOff':>7} {'delta_avg':>10} {'expected':>8}")
    print("-" * 65)

    pi = 1
    for (label, font, sz, rule, spacing) in tests:
        ys = []
        for i in range(N):
            p = doc.Paragraphs(pi)
            word.Selection.SetRange(p.Range.Start, p.Range.Start)
            y = float(word.Selection.Information(6))
            ys.append(y)

            if i == 0:
                # Check attributes
                r = p.Format.LineSpacingRule
                s = p.Format.LineSpacing
                g = p.Format.DisableLineHeightGrid
            pi += 1

        deltas = [ys[i+1] - ys[i] for i in range(len(ys)-1)]
        avg = sum(deltas) / len(deltas) if deltas else 0
        expected = spacing if spacing else "auto"
        grid_off = "True" if g else "False"
        rule_names = {0: "Single", 1: "1.5", 2: "Double", 3: "Exact", 4: "AtLst", 5: "Multi"}
        print(f"{label:<18} {rule_names.get(r,'?'):>6} {s:>8.2f} {grid_off:>7} {avg:>10.3f} {str(expected):>8}")

    doc.Close(False)
finally:
    word.Quit()

try:
    os.remove(test_path)
except:
    pass

print("\nDone.")
