"""Test if CJK fonts have a minimum line height that Exactly mode can't override.
Also verify grid settings are properly disabled."""
import win32com.client
import os
import time
import subprocess

subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

test_path = os.path.abspath("_test_min.docx")

# Test configs: (label, font, size, exactly_value)
configs = [
    # Meiryo tests - try to set Exactly below expected minimum
    ("Mei_Exact5", "Meiryo", 10.5, 5.0),
    ("Mei_Exact10", "Meiryo", 10.5, 10.0),
    ("Mei_Exact15", "Meiryo", 10.5, 15.0),
    ("Mei_Exact18", "Meiryo", 10.5, 18.0),
    ("Mei_Exact20", "Meiryo", 10.5, 20.0),
    ("Mei_Exact25", "Meiryo", 10.5, 25.0),
    ("Mei_Single", "Meiryo", 10.5, None),  # Single for comparison
    # Yu Gothic tests
    ("YG_Exact5", "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 5.0),
    ("YG_Exact15", "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 15.0),
    ("YG_Exact17", "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 17.0),
    ("YG_Exact18", "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 18.0),
    ("YG_Single", "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, None),
    # Calibri (western, should have NO minimum)
    ("Cal_Exact5", "Calibri", 10.5, 5.0),
    ("Cal_Exact10", "Calibri", 10.5, 10.0),
    ("Cal_Single", "Calibri", 10.5, None),
    # MS Gothic
    ("MSG_Exact5", "\uFF2D\uFF33 \u30B4\u30B7\u30C3\u30AF", 10.5, 5.0),
    ("MSG_Exact10", "\uFF2D\uFF33 \u30B4\u30B7\u30C3\u30AF", 10.5, 10.0),
    ("MSG_Exact13", "\uFF2D\uFF33 \u30B4\u30B7\u30C3\u30AF", 10.5, 13.0),
    ("MSG_Single", "\uFF2D\uFF33 \u30B4\u30B7\u30C3\u30AF", 10.5, None),
]

N = 5

print("Phase 1: Creating document...")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Add()
    time.sleep(0.5)

    # Set defaults
    ns = doc.Styles(-1)
    ns.Font.Name = "Calibri"
    ns.Font.Size = 10.5
    ns.ParagraphFormat.SpaceBefore = 0
    ns.ParagraphFormat.SpaceAfter = 0

    # Disable grid
    ps = doc.Sections(1).PageSetup
    ps.LayoutMode = 0  # wdLayoutModeDefault
    time.sleep(0.3)
    print(f"  LayoutMode after set: {ps.LayoutMode}")

    # Create paragraphs
    first = True
    for (label, font, sz, exact_val) in configs:
        for i in range(N):
            if not first:
                doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
            first = False
            pn = doc.Paragraphs.Count
            doc.Paragraphs(pn).Range.Text = f"{label} L{i+1}"

    time.sleep(0.3)

    # Format
    pi = 1
    for (label, font, sz, exact_val) in configs:
        for i in range(N):
            p = doc.Paragraphs(pi)
            p.Format.DisableLineHeightGrid = True
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Range.Font.Name = font
            p.Range.Font.Size = sz
            if exact_val is not None:
                p.Format.LineSpacingRule = 3  # wdLineSpaceExactly
                p.Format.LineSpacing = exact_val
            else:
                p.Format.LineSpacingRule = 0  # wdLineSpaceSingle
            pi += 1

    doc.SaveAs2(test_path)
    doc.Close(False)
    print(f"  Saved: {test_path}")
finally:
    word.Quit()
    time.sleep(2)

# Phase 2: Measure
print("\nPhase 2: Measuring...")
subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(test_path)
    time.sleep(0.5)

    # Check grid
    ps = doc.Sections(1).PageSetup
    print(f"  Re-opened LayoutMode: {ps.LayoutMode}")
    try:
        print(f"  LinePitch: {ps.LinePitch}")
    except:
        pass

    print(f"\n{'Label':<15} {'Rule':>5} {'SetVal':>6} {'Spacing':>8} {'Grid':>5} {'delta_avg':>10} {'minimum?':>8}")
    print("-" * 65)

    pi = 1
    for (label, font, sz, exact_val) in configs:
        ys = []
        for i in range(N):
            p = doc.Paragraphs(pi)
            word.Selection.SetRange(p.Range.Start, p.Range.Start)
            y = float(word.Selection.Information(6))
            ys.append(y)

            if i == 0:
                r = p.Format.LineSpacingRule
                s = p.Format.LineSpacing
                g = p.Format.DisableLineHeightGrid
            pi += 1

        deltas = [ys[j+1] - ys[j] for j in range(len(ys)-1)]
        avg = sum(deltas) / len(deltas) if deltas else 0
        rule_names = {0: "Singl", 1: "1.5", 2: "Dbl", 3: "Exact", 4: "AtLst", 5: "Multi"}
        set_str = f"{exact_val}" if exact_val else "auto"
        grid_str = "OFF" if g else "ON"
        has_min = "YES" if (exact_val and avg > exact_val + 0.5) else ""
        print(f"{label:<15} {rule_names.get(r,'?'):>5} {set_str:>6} {s:>8.2f} {grid_str:>5} {avg:>10.3f} {has_min:>8}")

    doc.Close(False)
finally:
    word.Quit()

try:
    os.remove(test_path)
except:
    pass

print("\nDone.")
