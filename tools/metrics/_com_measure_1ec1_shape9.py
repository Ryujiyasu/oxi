# -*- coding: utf-8 -*-
"""Measure Shape 9 P1 (BOX[5] = □３) position via Word COM API directly.

If COM reports 55.32pt (matches PDF), the property is still hidden in OOXML.
If COM reports ~47pt (matches my V_M0 clone), the gap is in PDF rendering pipeline."""
import sys, os, time
import pythoncom, win32com.client as wc
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")

# Word constants
wdActiveEndPageNumber = 3
wdHorizontalPositionRelativeToTextBoundary = 5  # Information(5)
wdHorizontalPositionRelativeToPage = 7  # Information(7)
wdShapePositionRelativeToTextBoundary = 4

pythoncom.CoInitialize()
word = None
for attempt in range(5):
    try:
        word = wc.Dispatch("Word.Application")
        time.sleep(2)
        word.Visible = False
        word.DisplayAlerts = False
        break
    except Exception as e:
        print(f"Word startup {attempt+1}: {e}")
        time.sleep(6)
if word is None:
    print("Failed Word"); sys.exit(1)

try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    print(f"Opened: {doc.Name}")

    # Iterate Shapes
    print(f"\nShape count: {doc.Shapes.Count}")
    print(f"InlineShape count: {doc.InlineShapes.Count}")

    for i in range(1, doc.Shapes.Count + 1):
        shape = doc.Shapes(i)
        try:
            name = shape.Name
        except: name = "?"
        try:
            tf = shape.TextFrame
            has_text = tf.HasText
            text = tf.TextRange.Text[:50] if has_text else ""
        except: text = ""
        try:
            left = shape.Left
            top = shape.Top
            width = shape.Width
            height = shape.Height
        except: left = top = width = height = None
        is_box = '□' in (text if text else '')
        marker = ' <-- HAS BOX' if is_box else ''
        print(f"\nShape {i}: name={name!r} text={text!r}{marker}")
        print(f"  Left={left} Top={top} W={width} H={height}")
        if is_box:
            # Examine first paragraph in textbox
            try:
                p1 = tf.TextRange.Paragraphs(1)
                pf = p1.Range.ParagraphFormat
                print(f"  P1 LeftIndent={pf.LeftIndent} FirstLineIndent={pf.FirstLineIndent}")
                print(f"  P1 SpaceBefore={pf.SpaceBefore} SpaceAfter={pf.SpaceAfter}")
                print(f"  P1 Text={p1.Range.Text[:50]!r}")
                # Position via Information API
                hp_text = p1.Range.Information(wdHorizontalPositionRelativeToTextBoundary)
                hp_page = p1.Range.Information(wdHorizontalPositionRelativeToPage)
                print(f"  P1 HPos relativeToTextBoundary (5) = {hp_text}pt")
                print(f"  P1 HPos relativeToPage (7) = {hp_page}pt")
            except Exception as e:
                print(f"  Para inspect failed: {e}")
    doc.Close(SaveChanges=False)
finally:
    try: word.Quit()
    except: pass
