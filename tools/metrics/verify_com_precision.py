"""Verify COM measurement precision by using Exactly spacing.
Also measure actual line heights via PDF export + PyMuPDF."""
import win32com.client
import os
import time
import subprocess

subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

test_path = os.path.abspath("_test_precision.docx")
pdf_path = os.path.abspath("_test_precision.pdf")

# Phase 1: Create test document
print("Phase 1: Creating test document...")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Add()
    time.sleep(0.5)

    ns = doc.Styles(-1)
    ns.Font.Name = "Calibri"
    ns.Font.Size = 10.5
    ns.ParagraphFormat.SpaceBefore = 0
    ns.ParagraphFormat.SpaceAfter = 0

    # Test A: Exactly spacing at known values (10 paras each)
    # Test B: Single spacing at various fonts (10 paras each)
    configs = [
        # (label, font, size, spacing_rule, spacing_val, count)
        ("Exact20", "Calibri", 10.5, 3, 20.0, 10),    # Exactly 20pt
        ("Exact15", "Calibri", 10.5, 3, 15.0, 10),    # Exactly 15pt
        ("SingleCal", "Calibri", 10.5, 0, None, 10),   # Single, Calibri
        ("SingleYuG", "\u6E38\u30B4\u30B7\u30C3\u30AF", 10.5, 0, None, 10),  # Single, Yu Gothic
        ("SingleMSG", "\uFF2D\uFF33 \u30B4\u30B7\u30C3\u30AF", 10.5, 0, None, 10),  # Single, MS Gothic
        ("SingleAri", "Arial", 10.5, 0, None, 10),     # Single, Arial
        ("SingleTNR", "Times New Roman", 10.5, 0, None, 10),
    ]

    first = True
    for (label, font, sz, rule, val, count) in configs:
        for i in range(count):
            if not first:
                doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
            first = False
            pn = doc.Paragraphs.Count
            p = doc.Paragraphs(pn)
            p.Range.Text = f"{label} L{i+1}"

    time.sleep(0.5)

    # Format
    pi = 1
    for (label, font, sz, rule, val, count) in configs:
        for i in range(count):
            p = doc.Paragraphs(pi)
            p.Format.DisableLineHeightGrid = True
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Format.LineSpacingRule = rule
            if val is not None:
                p.Format.LineSpacing = val
            p.Range.Font.Name = font
            p.Range.Font.Size = sz
            pi += 1

    doc.SaveAs2(test_path)

    # Export to PDF
    doc.ExportAsFixedFormat(pdf_path, 17)  # wdExportFormatPDF
    doc.Close(False)
    print(f"  Saved: {test_path}")
    print(f"  PDF: {pdf_path}")
    print(f"  Total paragraphs: {pi - 1}")
finally:
    word.Quit()
    time.sleep(2)

# Phase 2: Measure via COM
print("\nPhase 2: COM measurement...")
subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(3)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(test_path)
    time.sleep(0.5)

    pi = 1
    for (label, font, sz, rule, val, count) in configs:
        ys = []
        for i in range(count):
            word.Selection.SetRange(doc.Paragraphs(pi).Range.Start,
                                    doc.Paragraphs(pi).Range.Start)
            y = float(word.Selection.Information(6))
            ys.append(y)
            pi += 1

        # Compute deltas
        deltas = [ys[i+1] - ys[i] for i in range(len(ys)-1)]
        avg = sum(deltas) / len(deltas) if deltas else 0
        mn = min(deltas) if deltas else 0
        mx = max(deltas) if deltas else 0

        expected = val if val else "?"
        print(f"  {label:<12} expected={expected!s:<6} COM: avg={avg:.3f}  min={mn:.2f}  max={mx:.2f}  spread={mx-mn:.2f}")

    doc.Close(False)
finally:
    word.Quit()

# Phase 3: PDF measurement via PyMuPDF
print("\nPhase 3: PDF measurement (PyMuPDF)...")
try:
    import fitz  # PyMuPDF
    pdf_doc = fitz.open(pdf_path)
    page = pdf_doc[0]
    blocks = page.get_text("dict")["blocks"]

    # Extract Y positions of text lines
    y_positions = []
    for block in blocks:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    y = span["origin"][1]
                    text = span["text"][:20]
                    y_positions.append((y, text))

    # Group by config label
    pi = 0
    for (label, font, sz, rule, val, count) in configs:
        ys = []
        for i in range(count):
            if pi < len(y_positions):
                ys.append(y_positions[pi][0])
            pi += 1

        deltas = [ys[i+1] - ys[i] for i in range(len(ys)-1)]
        avg = sum(deltas) / len(deltas) if deltas else 0
        mn = min(deltas) if deltas else 0
        mx = max(deltas) if deltas else 0

        expected = val if val else "?"
        print(f"  {label:<12} expected={expected!s:<6} PDF: avg={avg:.3f}  min={mn:.2f}  max={mx:.2f}  spread={mx-mn:.2f}")

    pdf_doc.close()
except ImportError:
    print("  PyMuPDF not available - skipping PDF measurement")

# Cleanup
for f in [test_path, pdf_path]:
    try:
        os.remove(f)
    except:
        pass

print("\nDone.")
