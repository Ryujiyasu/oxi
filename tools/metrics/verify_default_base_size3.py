"""Create a test docx, then measure it separately. Two-phase approach."""
import win32com.client
import os
import time
import subprocess

subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

test_path = os.path.abspath("test_default_base.docx")

# Phase 1: Create test document
print("Phase 1: Creating test document...")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Add()
    time.sleep(1)

    # Set default font: Calibri 10.5pt (different from 11pt to discriminate)
    ns = doc.Styles(-1)
    ns.Font.Name = "Calibri"
    ns.Font.Size = 10.5
    ns.ParagraphFormat.LineSpacingRule = 0  # wdLineSpaceSingle
    ns.ParagraphFormat.SpaceBefore = 0
    ns.ParagraphFormat.SpaceAfter = 0

    # Disable grid for all content
    sec = doc.Sections(1)
    # Don't set LayoutMode to avoid instability

    # Create test paragraphs - pairs of identical content at different font/size
    test_configs = [
        ("Calibri", 10.5),
        ("Calibri", 14.0),
        ("Calibri", 20.0),
        ("游ゴシック", 10.5),
        ("游ゴシック", 14.0),
        ("游ゴシック", 20.0),
        ("Arial", 10.5),
        ("Arial", 14.0),
        ("Arial", 20.0),
        ("ＭＳ ゴシック", 10.5),
        ("ＭＳ ゴシック", 14.0),
        ("ＭＳ ゴシック", 20.0),
        ("Times New Roman", 10.5),
        ("Times New Roman", 14.0),
        ("Times New Roman", 24.0),
    ]

    # Build document: for each config, add 2 identical paragraphs
    # with DisableLineHeightGrid = True
    rng = doc.Range(0, 0)
    first = True
    for (font_name, font_size) in test_configs:
        if not first:
            rng.InsertAfter("\r")
            rng = doc.Range(rng.End - 1, rng.End)
        first = False

        # Para 1
        rng.InsertAfter(f"Test {font_name} {font_size}pt line1\r")
        # Para 2
        rng.InsertAfter(f"Test {font_name} {font_size}pt line2")

    time.sleep(0.5)

    # Now format each paragraph pair
    para_idx = 1
    for (font_name, font_size) in test_configs:
        for j in range(2):
            p = doc.Paragraphs(para_idx)
            p.Format.DisableLineHeightGrid = True
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Format.LineSpacingRule = 0  # Single
            p.Range.Font.Name = font_name
            p.Range.Font.Size = font_size
            para_idx += 1

    doc.SaveAs2(test_path)
    doc.Close(False)
    print(f"  Saved: {test_path}")
    print(f"  Total paragraphs: {para_idx - 1}")
finally:
    word.Quit()
    time.sleep(2)

# Phase 2: Measure
print("\nPhase 2: Measuring line heights...")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(test_path)
    time.sleep(1)

    total_paras = doc.Paragraphs.Count
    print(f"Total paragraphs: {total_paras}")

    # Get default font info
    ns = doc.Styles(-1)
    print(f"Default: {ns.Font.Name} {ns.Font.Size}pt")

    print(f"\n{'#':>3} {'Font':<20} {'Size':>5} {'Y':>8} {'Delta':>8}")
    print("-" * 50)

    positions = []
    for i in range(1, total_paras + 1):
        p = doc.Paragraphs(i)
        word.Selection.SetRange(p.Range.Start, p.Range.Start)
        y = float(word.Selection.Information(6))
        text = p.Range.Text[:40].replace('\r', '')
        positions.append((y, text))
        print(f"{i:>3} {text:<40} {y:>8.2f}")

    doc.Close(False)

    # Compute deltas for each pair
    print(f"\n{'Font':<20} {'Size':>5} {'COM_delta':>10}")
    print("-" * 40)

    for i in range(0, len(positions) - 1, 2):
        y1 = positions[i][0]
        y2 = positions[i+1][0]
        delta = y2 - y1
        text = positions[i][1]
        print(f"  {text[:35]:<35} delta={delta:.2f}")

finally:
    word.Quit()
    # Cleanup
    try:
        os.remove(test_path)
    except:
        pass
    print("\nDone.")
