"""Check para 2 details: text length, font size, line count."""
import win32com.client
import os
import time

docx_path = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\1ec1091177b1_006.docx"
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path)
    time.sleep(1)

    para2 = doc.Paragraphs(2)
    rng = para2.Range
    print(f"Para 2 text length: {len(rng.Text)} chars")
    print(f"Para 2 font size: {rng.Font.Size}pt")
    print(f"Para 2 font name: {rng.Font.Name}")
    print(f"Para 2 text (first 200): {rng.Text[:200]}")
    
    # Count lines in para 2
    word.Selection.SetRange(rng.Start, rng.End)
    line_count = word.Selection.Information(10)  # wdFirstCharacterLineNumber at end
    word.Selection.SetRange(rng.Start, rng.Start)
    line_start = word.Selection.Information(10)
    print(f"\nLine range: {line_start} to {line_count}")
    
    # Also check para 1
    para1 = doc.Paragraphs(1)
    rng1 = para1.Range
    print(f"\nPara 1 text: '{rng1.Text[:80]}'")
    print(f"Para 1 font size: {rng1.Font.Size}pt")
    print(f"Para 1 font name: {rng1.Font.Name}")

    # Check spacing
    pf2 = para2.Format
    print(f"\nPara 2 spacing: before={pf2.SpaceBefore}pt, after={pf2.SpaceAfter}pt")
    print(f"Para 2 line spacing: {pf2.LineSpacing}pt (rule={pf2.LineSpacingRule})")

    pf1 = para1.Format
    print(f"Para 1 spacing: before={pf1.SpaceBefore}pt, after={pf1.SpaceAfter}pt")
    print(f"Para 1 line spacing: {pf1.LineSpacing}pt (rule={pf1.LineSpacingRule})")
    
    doc.Close(False)
finally:
    word.Quit()
