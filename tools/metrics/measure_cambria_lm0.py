"""Measure Cambria LM=0 line heights for lm0_lineauto.json.

Creates a temp doc with Cambria text at each font size, no docGrid (LM=0),
and measures the Y gap between consecutive paragraphs.
"""
import win32com.client, time, sys, os, json
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

doc = word.Documents.Add()
time.sleep(0.3)

# Remove docGrid (ensure LM=0)
doc.PageSetup.LayoutMode = 0  # wdLayoutModeDefault = no grid

# Generate test: two paragraphs at each font size
sizes = [round(7.0 + i * 0.5, 1) for i in range(37)]  # 7.0 to 25.0

results = {}
for fs in sizes:
    # Clear document
    doc.Content.Delete()
    time.sleep(0.1)

    # Add two paragraphs with Cambria at this size
    rng = doc.Content
    rng.Text = "A\rA\r"
    rng.Font.Name = "Cambria"
    rng.Font.Size = fs
    # Single spacing, no space before/after
    rng.ParagraphFormat.LineSpacingRule = 0  # wdLineSpaceSingle
    rng.ParagraphFormat.SpaceBefore = 0
    rng.ParagraphFormat.SpaceAfter = 0

    time.sleep(0.1)

    y1 = doc.Paragraphs(1).Range.Information(6)
    y2 = doc.Paragraphs(2).Range.Information(6)
    lh = round(y2 - y1, 1)
    results[str(fs)] = lh
    print(f"  Cambria {fs}pt: line_height = {lh}pt")

doc.Close(SaveChanges=False)
word.Quit()

print(f"\nResults:")
print(json.dumps(results, indent=2))
