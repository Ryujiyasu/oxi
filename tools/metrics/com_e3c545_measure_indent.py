"""COM-measure Word's actual indent for e3c545 w_i=30 paragraph.

Hypothesis: Word uses leftChars × paragraph_font_size (10.5pt indent)
while Oxi uses left twip value (12pt indent). Difference = 1.5pt
causes Oxi to wrap "い" to line 2 (overflow 0.6pt) while Word fits it.
"""
import os
import sys
import win32com.client

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.normpath(os.path.join(REPO, "tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx"))

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
try:
    # word_i=30: "メタデータは、各機関で独自に定義します。具体例は、..."
    # Get the paragraph
    p = doc.Paragraphs(30)
    pf = p.Format
    print(f"para[30] indent properties:")
    print(f"  LeftIndent (pt): {pf.LeftIndent:.3f}")
    print(f"  RightIndent (pt): {pf.RightIndent:.3f}")
    print(f"  FirstLineIndent (pt): {pf.FirstLineIndent:.3f}")
    print(f"  CharacterUnitLeftIndent: {pf.CharacterUnitLeftIndent:.3f}")
    print(f"  CharacterUnitRightIndent: {pf.CharacterUnitRightIndent:.3f}")
    print(f"  CharacterUnitFirstLineIndent: {pf.CharacterUnitFirstLineIndent:.3f}")

    # Also measure the X position of the first run's first char
    r = p.Range
    start_r = doc.Range(r.Start, r.Start)
    wdHorizontalPositionRelativeToPage = 5
    wdVerticalPositionRelativeToPage = 6
    x = start_r.Information(wdHorizontalPositionRelativeToPage)
    y = start_r.Information(wdVerticalPositionRelativeToPage)
    print(f"\n  First char Y position: {y:.2f}pt")
    print(f"  First char X position: {x:.2f}pt")

    # Body margin info
    sec = doc.Sections(1)
    ps = sec.PageSetup
    print(f"\n  Section margins: L={ps.LeftMargin:.1f} R={ps.RightMargin:.1f}")
    print(f"  Section page width: {ps.PageWidth:.1f}")
    body_left = ps.LeftMargin
    body_right = ps.PageWidth - ps.RightMargin
    print(f"  Body left: {body_left:.2f}, body right: {body_right:.2f}")
    print(f"  Body width: {body_right - body_left:.2f}")

    # Compare with Oxi
    print(f"\n  Oxi uses: 12pt left indent (from w:left=240 twips)")
    print(f"  Word reports LeftIndent: {pf.LeftIndent:.2f}pt")
    print(f"  Difference: {12.0 - pf.LeftIndent:.2f}pt (Oxi extra)")

    # Word's actual font size for this paragraph
    # Check first run
    fs = p.Range.Font.Size
    print(f"\n  Paragraph font size (Word): {fs}pt")
finally:
    doc.Close(SaveChanges=False)
    word.Quit()
