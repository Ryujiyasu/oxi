"""COM: Measure body paragraph available width and line breaks.

For kyodokenkyuyoushiki01 page 3 paragraphs (P196-P202 in Oxi),
measure exact line breaks, character widths, and available width.
"""
import win32com.client
import os, time, json


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    path = os.path.abspath("tools/golden-test/documents/docx/459f05f1e877_kyodokenkyuyoushiki01.docx")
    doc = word.Documents.Open(path, ReadOnly=True)
    time.sleep(1)

    # Page setup
    sec = doc.Sections(1)
    ps = sec.PageSetup
    print(f"Page: {ps.PageWidth:.1f}pt x {ps.PageHeight:.1f}pt")
    print(f"Margins: L={ps.LeftMargin:.1f} R={ps.RightMargin:.1f} T={ps.TopMargin:.1f} B={ps.BottomMargin:.1f}")
    print(f"Content width: {ps.PageWidth - ps.LeftMargin - ps.RightMargin:.1f}pt")
    print(f"CharsLine: {ps.CharsLine}, LinesPage: {ps.LinesPage}")

    # Grid settings
    try:
        print(f"LayoutMode: {sec.PageSetup.LayoutMode}")
    except:
        pass

    # Find paragraphs on page 3
    print(f"\n=== Paragraphs on page 3 ===")
    total_paras = doc.Paragraphs.Count
    print(f"Total paragraphs: {total_paras}")

    # Find the body paragraphs after the tables
    for pi in range(1, total_paras + 1):
        para = doc.Paragraphs(pi)
        rng = para.Range
        page = rng.Information(3)  # wdActiveEndPageNumber

        if page < 3:
            continue
        if page > 3:
            break

        text = rng.Text.rstrip('\r\n')
        if not text:
            continue

        y = rng.Information(6)  # wdVerticalPositionRelativeToPage
        x = rng.Information(5)  # wdHorizontalPositionRelativeToPage

        # Line count via Selection approach
        # Instead, count characters and check line breaks via horizontal positions
        chars = rng.Characters
        n = chars.Count

        # Get first and last char positions
        first_x = chars(1).Information(5)
        first_y = chars(1).Information(6)

        # Count lines by checking Y positions of every Nth character
        lines = []
        current_line_y = None
        current_line_chars = 0
        current_line_start_x = None

        step = max(1, n // 200)  # sample every Nth char for speed
        for i in range(1, n + 1, step):
            try:
                ch_y = round(chars(i).Information(6), 1)
                ch_x = chars(i).Information(5)

                if current_line_y is None or abs(ch_y - current_line_y) > 1.0:
                    if current_line_y is not None:
                        lines.append({
                            'y': current_line_y,
                            'chars': current_line_chars,
                            'start_x': current_line_start_x,
                        })
                    current_line_y = ch_y
                    current_line_chars = 0
                    current_line_start_x = ch_x

                current_line_chars += step
            except:
                pass

        if current_line_y is not None:
            lines.append({
                'y': current_line_y,
                'chars': current_line_chars,
                'start_x': current_line_start_x,
            })

        # Font info
        font_name = rng.Font.Name
        font_size = rng.Font.Size

        # Paragraph format
        left_indent = para.Format.LeftIndent
        right_indent = para.Format.RightIndent
        first_line_indent = para.Format.FirstLineIndent

        print(f"\nP{pi} page={page} y={y:.1f} x={x:.1f}")
        print(f"  font={font_name} size={font_size}")
        print(f"  indent: L={left_indent:.1f} R={right_indent:.1f} FI={first_line_indent:.1f}")
        print(f"  total_chars={len(text)} lines_detected={len(lines)}")
        if lines:
            print(f"  line1: y={lines[0]['y']:.1f} chars~{lines[0]['chars']} start_x={lines[0]['start_x']:.1f}")
            if len(lines) > 1:
                print(f"  line2: y={lines[1]['y']:.1f} chars~{lines[1]['chars']} start_x={lines[1]['start_x']:.1f}")
        print(f"  text: \"{text[:80]}\"")

        if pi > 25 and page >= 3:
            break

    doc.Close(SaveChanges=False)
    word.Quit()


if __name__ == "__main__":
    main()
