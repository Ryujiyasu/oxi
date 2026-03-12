"""
Measure specific properties of comprehensive_test.docx via Word COM.
Targets:
1. Mixed font size line heights
2. Heading 1 space_before
3. CJK+Latin mixed line height
4. Table border width
"""
import win32com.client
import os, time

DOCX_PATH = os.path.abspath(r"tests/fixtures/comprehensive_test.docx")

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Open(DOCX_PATH)
        time.sleep(2)
        doc.Repaginate()
        time.sleep(1)

        n_paras = doc.Paragraphs.Count
        print(f"Total paragraphs: {n_paras}")
        print()

        # Measure all paragraph positions and text
        print("=== All Paragraph Positions ===")
        print(f"{'#':>3} {'Y pos':>8} {'Delta':>8} {'Text (first 60 chars)'}")
        print("-" * 90)

        prev_y = None
        for i in range(1, min(n_paras + 1, 60)):  # First 60 paragraphs
            para = doc.Paragraphs(i)
            rng = para.Range
            text = rng.Text.strip()[:60]
            rng_start = para.Range
            rng_start.Collapse(1)  # wdCollapseStart

            y = rng_start.Information(6)  # wdVerticalPositionRelativeToPage
            page = rng_start.Information(3)  # wdActiveEndPageNumber

            delta = y - prev_y if prev_y is not None else 0
            prev_y = y

            # Check if this is a heading
            style_name = para.Style.NameLocal
            marker = ""
            if "Heading" in style_name or "heading" in style_name.lower():
                marker = f" [{style_name}]"

            print(f"{i:3d} {y:8.2f} {delta:8.2f} {text}{marker}")

            # Reset prev_y on page break
            if i > 1 and y < 100:
                prev_y = y  # new page

        # Specifically measure Heading 1
        print("\n=== Heading Style Details ===")
        for i in range(1, min(n_paras + 1, 60)):
            para = doc.Paragraphs(i)
            style = para.Style
            if "Heading" in style.NameLocal:
                fmt = para.Format
                print(f"Para {i}: {style.NameLocal}")
                print(f"  SpaceBefore: {fmt.SpaceBefore}pt")
                print(f"  SpaceAfter: {fmt.SpaceAfter}pt")
                print(f"  LineSpacing: {fmt.LineSpacing}pt")
                print(f"  LineSpacingRule: {fmt.LineSpacingRule}")
                # Get font info
                rng = para.Range
                print(f"  Font: {rng.Font.Name}, Size: {rng.Font.Size}pt")

        # Measure CJK paragraphs specifically
        print("\n=== CJK/Mixed Paragraphs ===")
        for i in range(1, min(n_paras + 1, 60)):
            para = doc.Paragraphs(i)
            text = para.Range.Text.strip()
            if any(ord(c) > 0x3000 for c in text):
                rng = para.Range
                rng.Collapse(1)
                y = rng.Information(6)
                fmt = para.Format
                print(f"Para {i}: Y={y:.2f}pt")
                print(f"  Text: {text[:80]}")
                print(f"  LineSpacing: {fmt.LineSpacing}pt, Rule: {fmt.LineSpacingRule}")
                print(f"  Font: {rng.Font.Name}, Size: {rng.Font.Size}pt")
                # Check east asian font
                print(f"  EastAsiaFont: {rng.Font.NameFarEast}")

        # Table border measurement
        print("\n=== Table Borders ===")
        n_tables = doc.Tables.Count
        print(f"Total tables: {n_tables}")
        for t in range(1, min(n_tables + 1, 4)):
            table = doc.Tables(t)
            borders = table.Borders
            print(f"\nTable {t}:")
            # Border types: -1=Top, -2=Left, -3=Bottom, -4=Right, -5=Horizontal, -6=Vertical
            border_names = {-1: "Top", -2: "Left", -3: "Bottom", -4: "Right", -5: "InsideH", -6: "InsideV"}
            for bid, bname in border_names.items():
                try:
                    b = borders(bid)
                    # LineWidth is in 8ths of a point
                    print(f"  {bname}: Width={b.LineWidth/8:.1f}pt, Style={b.LineStyle}, Color={b.Color}")
                except Exception as e:
                    print(f"  {bname}: {e}")

        # Document grid info
        print("\n=== Document Grid ===")
        section = doc.Sections(1)
        pf = section.PageSetup
        print(f"  TopMargin: {pf.TopMargin}pt")
        print(f"  BottomMargin: {pf.BottomMargin}pt")
        print(f"  LeftMargin: {pf.LeftMargin}pt")
        print(f"  RightMargin: {pf.RightMargin}pt")
        print(f"  LinePitch (grid): {pf.LinesPage}")

        # Default font
        print(f"\n=== Default Font ===")
        default_style = doc.Styles("Normal")
        print(f"  Font: {default_style.Font.Name}")
        print(f"  Size: {default_style.Font.Size}pt")
        print(f"  EastAsia: {default_style.Font.NameFarEast}")

        doc.Close(0)
    finally:
        word.Quit()

if __name__ == "__main__":
    main()
