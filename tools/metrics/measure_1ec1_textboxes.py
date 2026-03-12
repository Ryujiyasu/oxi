"""Measure TextBox positions from 1ec1091177b1_006.docx using Word COM automation."""
import win32com.client
import os
import sys

doc_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'golden-test', 'documents', 'docx', '1ec1091177b1_006.docx'))
print(f"Document: {doc_path}")
print(f"Exists: {os.path.exists(doc_path)}")
print()

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

try:
    doc = word.Documents.Open(doc_path)

    # Page setup
    sec = doc.Sections(1)
    print("=== PAGE SETUP ===")
    print(f"PageWidth={sec.PageSetup.PageWidth}pt")
    print(f"PageHeight={sec.PageSetup.PageHeight}pt")
    print(f"LeftMargin={sec.PageSetup.LeftMargin}pt")
    print(f"RightMargin={sec.PageSetup.RightMargin}pt")
    print(f"TopMargin={sec.PageSetup.TopMargin}pt")
    print(f"BottomMargin={sec.PageSetup.BottomMargin}pt")
    left_margin = sec.PageSetup.LeftMargin
    top_margin = sec.PageSetup.TopMargin
    print()

    # RelativeHorizontalPosition constants
    rh_names = {0: 'margin', 1: 'page', 2: 'column', 3: 'character',
                4: 'leftMargin', 5: 'rightMargin', 6: 'insideMargin', 7: 'outsideMargin'}
    # RelativeVerticalPosition constants
    rv_names = {0: 'margin', 1: 'page', 2: 'paragraph', 3: 'line'}
    # WrapFormat.Type constants
    wrap_names = {0: 'inline', 1: 'topBottom', 2: 'square', 3: 'none', 4: 'tight', 5: 'through', 6: 'unknown'}

    # Shapes (floating)
    print(f"=== SHAPES (Count={doc.Shapes.Count}) ===")
    for i in range(1, doc.Shapes.Count + 1):
        shape = doc.Shapes(i)
        print(f"Shape {i}: Name={shape.Name}")
        print(f"  Type={shape.Type}")
        print(f"  Left={shape.Left}pt, Top={shape.Top}pt")
        print(f"  Width={shape.Width}pt, Height={shape.Height}pt")

        rh = shape.RelativeHorizontalPosition
        rv = shape.RelativeVerticalPosition
        print(f"  RelativeHorizontalPosition={rh} ({rh_names.get(rh, '?')})")
        print(f"  RelativeVerticalPosition={rv} ({rv_names.get(rv, '?')})")

        wt = shape.WrapFormat.Type
        print(f"  WrapFormat.Type={wt} ({wrap_names.get(wt, '?')})")
        print(f"  WrapFormat.DistanceTop={shape.WrapFormat.DistanceTop}pt")
        print(f"  WrapFormat.DistanceBottom={shape.WrapFormat.DistanceBottom}pt")
        print(f"  WrapFormat.DistanceLeft={shape.WrapFormat.DistanceLeft}pt")
        print(f"  WrapFormat.DistanceRight={shape.WrapFormat.DistanceRight}pt")

        try:
            if shape.TextFrame.HasText:
                text_preview = shape.TextFrame.TextRange.Text[:80].replace('\r', '\\r').replace('\n', '\\n')
                print(f"  Text: \"{text_preview}\"")
            print(f"  TextFrame.MarginLeft={shape.TextFrame.MarginLeft}pt")
            print(f"  TextFrame.MarginRight={shape.TextFrame.MarginRight}pt")
            print(f"  TextFrame.MarginTop={shape.TextFrame.MarginTop}pt")
            print(f"  TextFrame.MarginBottom={shape.TextFrame.MarginBottom}pt")
        except Exception as e:
            print(f"  TextFrame error: {e}")

        # Compute absolute page position
        if rh == 0:  # margin-relative
            abs_x = left_margin + shape.Left
        elif rh == 1:  # page-relative
            abs_x = shape.Left
        elif rh == 2:  # column-relative (same as margin for single-column)
            abs_x = left_margin + shape.Left
        else:
            abs_x = shape.Left  # fallback

        if rv == 0:  # margin-relative
            abs_y = top_margin + shape.Top
        elif rv == 1:  # page-relative
            abs_y = shape.Top
        elif rv == 2:  # paragraph-relative
            abs_y = shape.Top  # can't compute without paragraph position
        else:
            abs_y = shape.Top

        print(f"  >> Absolute page pos: x={abs_x:.2f}pt, y={abs_y:.2f}pt")
        print()

    # InlineShapes
    print(f"=== INLINE SHAPES (Count={doc.InlineShapes.Count}) ===")
    for i in range(1, doc.InlineShapes.Count + 1):
        ishape = doc.InlineShapes(i)
        print(f"InlineShape {i}: Type={ishape.Type}, Width={ishape.Width}pt, Height={ishape.Height}pt")

    # Also check anchored shapes via Sections/Headers/Footers
    print()
    print("=== HEADER/FOOTER SHAPES ===")
    for si in range(1, doc.Sections.Count + 1):
        sec = doc.Sections(si)
        for hf_type, hf_name in [(1, 'PrimaryHeader'), (2, 'PrimaryFooter'), (3, 'FirstPageHeader'), (4, 'FirstPageFooter')]:
            try:
                hf = sec.Headers(hf_type) if hf_type <= 2 else sec.Footers(hf_type - 2)
                if hf.Shapes.Count > 0:
                    print(f"Section {si} {hf_name}: {hf.Shapes.Count} shapes")
                    for j in range(1, hf.Shapes.Count + 1):
                        s = hf.Shapes(j)
                        print(f"  Shape {j}: Left={s.Left}pt, Top={s.Top}pt, Width={s.Width}pt, Height={s.Height}pt")
            except:
                pass

    doc.Close(False)
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
finally:
    word.Quit()

print("\nDone.")
