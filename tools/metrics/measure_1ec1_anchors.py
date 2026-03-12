"""Get anchor paragraph info and check for page-2 shapes in 1ec1 document."""
import win32com.client
import os

doc_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'golden-test', 'documents', 'docx', '1ec1091177b1_006.docx'))

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

try:
    doc = word.Documents.Open(doc_path)

    print("=== DOCUMENT INFO ===")
    print(f"Pages: {doc.ComputeStatistics(2)}")  # wdStatisticPages=2
    print(f"Paragraphs: {doc.Paragraphs.Count}")
    print()

    # Get anchor paragraph for each shape
    print("=== SHAPE ANCHORS ===")
    for i in range(1, doc.Shapes.Count + 1):
        shape = doc.Shapes(i)
        anchor = shape.Anchor
        # Find which paragraph the anchor is in
        anchor_para = anchor.Paragraphs(1)
        anchor_text = anchor_para.Range.Text[:60].replace('\r', '\\r').replace('\n', '\\n')

        # Get the page number of the anchor
        anchor_page = anchor.Information(3)  # wdActiveEndPageNumber=3

        print(f"Shape {i} ({shape.Name}):")
        print(f"  Anchor page: {anchor_page}")
        print(f"  Anchor para text: \"{anchor_text}\"")
        print(f"  Left={shape.Left}pt, Top={shape.Top}pt")
        print(f"  RelH={shape.RelativeHorizontalPosition}, RelV={shape.RelativeVerticalPosition}")

        # For paragraph-relative shapes, try to get the anchor paragraph's position on page
        try:
            # Get position of anchor paragraph start
            para_range = anchor_para.Range
            para_left = para_range.Information(5)   # wdHorizontalPositionRelativeToPage
            para_top = para_range.Information(6)     # wdVerticalPositionRelativeToPage
            print(f"  Anchor para position on page: left={para_left}pt, top={para_top}pt")

            if shape.RelativeVerticalPosition == 2:  # paragraph-relative
                abs_y = para_top + shape.Top
                print(f"  >> True absolute Y = {para_top} + {shape.Top} = {abs_y}pt")
            if shape.RelativeHorizontalPosition == 0:  # margin-relative
                abs_x = 42.55 + shape.Left  # left margin
                print(f"  >> True absolute X = 42.55 + {shape.Left} = {abs_x}pt")
        except Exception as e:
            print(f"  Position query error: {e}")
        print()

    # Check Left=-999995 mystery - these might be on page 2
    print("=== SHAPES WITH EXTREME LEFT VALUES ===")
    print("These shapes (Left~=-999995) are likely on page 2.")
    print("In Word, Left=-999995 means 'wdShapeCenter' or hidden positioning.")
    print()

    # Try to get actual rendered positions via Range.Information
    print("=== PARAGRAPH POSITIONS (all paragraphs) ===")
    for i in range(1, min(doc.Paragraphs.Count + 1, 30)):
        para = doc.Paragraphs(i)
        r = para.Range
        page = r.Information(3)
        text = r.Text[:40].replace('\r', '\\r').replace('\n', '\\n')
        try:
            top = r.Information(6)
            left = r.Information(5)
            print(f"Para {i} (page {page}): top={top}pt, left={left}pt, text=\"{text}\"")
        except:
            print(f"Para {i} (page {page}): text=\"{text}\" [position N/A]")

    doc.Close(False)
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
finally:
    word.Quit()

print("\nDone.")
