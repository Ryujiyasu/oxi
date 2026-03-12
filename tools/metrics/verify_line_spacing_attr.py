"""Check what LineSpacing value COM reports for different settings."""
import win32com.client
import os
import time
import subprocess

subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

test_path = os.path.abspath("_test_attr.docx")

print("Phase 1: Creating document...")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Add()
    time.sleep(0.5)

    # Check default Normal style BEFORE any modification
    ns = doc.Styles(-1)
    print(f"  Default Normal: Font={ns.Font.Name} Size={ns.Font.Size}")
    print(f"  Default Normal spacing: Rule={ns.ParagraphFormat.LineSpacingRule} Spacing={ns.ParagraphFormat.LineSpacing}")
    print(f"  Default Normal SpaceBefore={ns.ParagraphFormat.SpaceBefore} SpaceAfter={ns.ParagraphFormat.SpaceAfter}")

    # Now set Normal to Calibri 10.5
    ns.Font.Name = "Calibri"
    ns.Font.Size = 10.5
    ns.ParagraphFormat.SpaceBefore = 0
    ns.ParagraphFormat.SpaceAfter = 0

    print(f"  After setting Calibri 10.5: Rule={ns.ParagraphFormat.LineSpacingRule} Spacing={ns.ParagraphFormat.LineSpacing}")

    # Test configs: (label, font, size, set_rule, set_spacing)
    configs = [
        ("Default",    "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, None, None),     # No LineSpacing modification
        ("Single",     "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 0, None),         # wdLineSpaceSingle
        ("Exact16.5",  "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 3, 16.5),         # Exactly = GDI@96 value
        ("Multiple1",  "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 5, 12.0),         # Multiple 1.0 (12pt = 240/240 * 12pt?)
        ("CalDefault", "Calibri", 10.5, None, None),
        ("CalSingle",  "Calibri", 10.5, 0, None),
        ("CalExact",   "Calibri", 10.5, 3, 12.75),
    ]

    # Create 5 paragraphs per config
    N = 5
    first = True
    for (label, font, sz, rule, spacing) in configs:
        for i in range(N):
            if not first:
                doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
            first = False
            pn = doc.Paragraphs.Count
            doc.Paragraphs(pn).Range.Text = f"{label} L{i+1}"

    time.sleep(0.3)

    # Format
    pi = 1
    for (label, font, sz, rule, spacing) in configs:
        for i in range(N):
            p = doc.Paragraphs(pi)
            p.Format.DisableLineHeightGrid = True
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Range.Font.Name = font
            p.Range.Font.Size = sz
            if rule is not None:
                p.Format.LineSpacingRule = rule
            if spacing is not None:
                p.Format.LineSpacing = spacing
            pi += 1

    doc.SaveAs2(test_path)
    doc.Close(False)
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

    print(f"\n{'Label':<14} {'Rule':>5} {'Spacing':>8} {'COM_delta':>10}")
    print("-" * 45)

    pi = 1
    for (label, font, sz, rule, spacing) in configs:
        ys = []
        reported_rules = []
        reported_spacings = []
        for i in range(N):
            p = doc.Paragraphs(pi)
            word.Selection.SetRange(p.Range.Start, p.Range.Start)
            y = float(word.Selection.Information(6))
            ys.append(y)
            reported_rules.append(p.Format.LineSpacingRule)
            reported_spacings.append(p.Format.LineSpacing)
            pi += 1

        deltas = [ys[i+1] - ys[i] for i in range(len(ys)-1)]
        avg_delta = sum(deltas) / len(deltas) if deltas else 0

        # Report first paragraph's attributes
        r = reported_rules[0]
        s = reported_spacings[0]
        rule_names = {0: "Single", 1: "1.5", 2: "Double", 3: "Exact", 4: "AtLeast", 5: "Multi"}
        rname = rule_names.get(r, str(r))
        print(f"{label:<14} {rname:>5} {s:>8.2f} {avg_delta:>10.3f}")

    doc.Close(False)
finally:
    word.Quit()

try:
    os.remove(test_path)
except:
    pass

print("\nDone.")
