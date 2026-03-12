"""Check horizontal alignment enum values for shapes with Left=-999995/-999996."""
import win32com.client
import os

doc_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'golden-test', 'documents', 'docx', '1ec1091177b1_006.docx'))

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

try:
    doc = word.Documents.Open(doc_path)
    left_margin = doc.Sections(1).PageSetup.LeftMargin
    right_margin = doc.Sections(1).PageSetup.RightMargin
    page_width = doc.Sections(1).PageSetup.PageWidth

    # -999995 = msoAlignCenter (some interpretations), let's check
    # Word uses special negative values as alignment codes:
    # -999995 = wdShapeCenter
    # -999996 = wdShapeRight
    # -999997 = wdShapeLeft
    # -999998 = wdShapeInside
    # -999999 = wdShapeOutside

    print("=== SHAPE HORIZONTAL ALIGNMENT ANALYSIS ===")
    print(f"Page: width={page_width}pt, leftMargin={left_margin}pt, rightMargin={right_margin}pt")
    print(f"Printable width = {page_width - left_margin - right_margin}pt")
    print()

    for i in range(1, doc.Shapes.Count + 1):
        shape = doc.Shapes(i)
        left_val = shape.Left

        if left_val < -900000:
            # This is an alignment constant, not a position
            align_map = {
                -999995: 'CENTER',
                -999996: 'RIGHT',
                -999997: 'LEFT',
                -999998: 'INSIDE',
                -999999: 'OUTSIDE',
            }
            align_name = align_map.get(int(left_val), f'UNKNOWN({left_val})')

            # Compute actual position based on alignment
            margin_width = page_width - left_margin - right_margin
            if int(left_val) == -999995:  # CENTER relative to margin
                actual_left = (margin_width - shape.Width) / 2
                abs_x = left_margin + actual_left
            elif int(left_val) == -999996:  # RIGHT relative to margin
                actual_left = margin_width - shape.Width
                abs_x = left_margin + actual_left
            elif int(left_val) == -999997:  # LEFT
                actual_left = 0
                abs_x = left_margin
            else:
                actual_left = 0
                abs_x = left_margin

            print(f"Shape {i} ({shape.Name}):")
            print(f"  Left={left_val} => {align_name}")
            print(f"  Width={shape.Width}pt, Height={shape.Height}pt")
            print(f"  RelH={shape.RelativeHorizontalPosition}")
            print(f"  Margin-relative left = {actual_left:.2f}pt")
            print(f"  Absolute X on page = {abs_x:.2f}pt")

            # Anchor para
            anchor = shape.Anchor
            anchor_page = anchor.Information(3)
            anchor_top = anchor.Information(6)
            print(f"  Anchor page={anchor_page}, anchor_top={anchor_top}pt")
            print(f"  Shape Top={shape.Top}pt (paragraph-relative)")
            print(f"  Absolute Y = {anchor_top + shape.Top:.2f}pt")

            try:
                if shape.TextFrame.HasText:
                    text = shape.TextFrame.TextRange.Text[:80].replace('\r', '\\r')
                    print(f"  Text: \"{text}\"")
                print(f"  TextFrame margins: L={shape.TextFrame.MarginLeft}, R={shape.TextFrame.MarginRight}, T={shape.TextFrame.MarginTop}, B={shape.TextFrame.MarginBottom}")
            except:
                pass
            print()

    # Summary table
    print("=== COMPLETE POSITION SUMMARY (all shapes, absolute page coordinates) ===")
    margin_width = page_width - left_margin - right_margin
    for i in range(1, doc.Shapes.Count + 1):
        shape = doc.Shapes(i)
        left_val = shape.Left
        anchor_top = shape.Anchor.Information(6)

        # X position
        if int(left_val) == -999995:  # CENTER
            rel_left = (margin_width - shape.Width) / 2
            abs_x = left_margin + rel_left
            h_note = "CENTER"
        elif int(left_val) == -999996:  # RIGHT
            rel_left = margin_width - shape.Width
            abs_x = left_margin + rel_left
            h_note = "RIGHT"
        elif shape.RelativeHorizontalPosition == 0:  # margin
            abs_x = left_margin + left_val
            h_note = f"margin+{left_val:.2f}"
        elif shape.RelativeHorizontalPosition == 4:  # leftMargin area
            abs_x = left_val
            h_note = f"leftMargin={left_val:.2f}"
        else:
            abs_x = left_val
            h_note = f"raw={left_val:.2f}"

        # Y position (all are paragraph-relative)
        abs_y = anchor_top + shape.Top

        name = shape.Name.encode('ascii', 'replace').decode()
        try:
            text = shape.TextFrame.TextRange.Text[:30].replace('\r', '\\r').encode('ascii', 'replace').decode()
        except:
            text = "[no text]"

        print(f"Shape {i}: abs=({abs_x:.2f}, {abs_y:.2f}) size=({shape.Width:.2f} x {shape.Height:.2f}) [{h_note}] \"{text}\"")

    # Also get border info
    print()
    print("=== SHAPE LINE/BORDER INFO ===")
    for i in range(1, doc.Shapes.Count + 1):
        shape = doc.Shapes(i)
        try:
            line = shape.Line
            print(f"Shape {i}: Visible={line.Visible}, Weight={line.Weight}pt, Style={line.Style}")
            fill = shape.Fill
            print(f"  Fill: Visible={fill.Visible}, Type={fill.Type}")
        except Exception as e:
            print(f"Shape {i}: Line/Fill error: {e}")

    doc.Close(False)
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
finally:
    word.Quit()

print("\nDone.")
