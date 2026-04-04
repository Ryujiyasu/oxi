"""COM: Measure exact cell available width for text wrapping.

For each table cell, measure:
- Cell width (Table.Cell.Width)
- Column width
- First char X position (left edge of text)
- Cell left border X
- Last char end X on last line
These determine the actual available width for wrapping.
"""
import win32com.client
import os, time, json


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    path = os.path.abspath("tools/golden-test/documents/docx/459f05f1e877_kyodokenkyuyoushiki01.docx")
    doc = word.Documents.Open(path, ReadOnly=True)
    time.sleep(1)

    # Focus on table 1
    tbl = doc.Tables(1)
    rows = tbl.Rows.Count
    cols = tbl.Columns.Count
    print(f"Table 1: {rows} rows x {cols} cols")

    # Table margins/padding
    print(f"\nTable-level settings:")
    try:
        print(f"  LeftPadding: {tbl.LeftPadding:.2f}pt")
        print(f"  RightPadding: {tbl.RightPadding:.2f}pt")
        print(f"  TopPadding: {tbl.TopPadding:.2f}pt")
        print(f"  BottomPadding: {tbl.BottomPadding:.2f}pt")
    except:
        pass

    results = []
    for r in range(1, min(rows + 1, 11)):
        for c in range(1, cols + 1):
            try:
                cell = tbl.Cell(r, c)
                rng = cell.Range
                text = rng.Text.rstrip('\r\x07')
                if not text:
                    continue

                cell_w = cell.Width

                # Cell padding
                try:
                    lp = cell.LeftPadding if hasattr(cell, 'LeftPadding') else -1
                    rp = cell.RightPadding if hasattr(cell, 'RightPadding') else -1
                except:
                    lp = rp = -1

                # First char position
                first_ch = rng.Characters(1)
                first_x = first_ch.Information(5)  # wdHorizontalPositionRelativeToPage

                # Cell range start position
                cell_start = rng.Duplicate
                cell_start.Collapse(1)  # start
                cell_start_x = cell_start.Information(5)

                # Paragraph indent
                para = rng.Paragraphs(1)
                left_indent = para.LeftIndent
                right_indent = para.RightIndent
                first_line_indent = para.FirstLineIndent

                # Character spacing
                cs = rng.Font.Spacing

                # charGrid pitch info
                try:
                    sec = rng.Sections(1)
                    char_pitch = sec.PageSetup.CharsLine
                    line_pitch = sec.PageSetup.LinesPage
                except:
                    char_pitch = line_pitch = -1

                entry = {
                    'row': r, 'col': c,
                    'cell_w': round(cell_w, 2),
                    'cell_lp': round(lp, 2) if lp >= 0 else None,
                    'cell_rp': round(rp, 2) if rp >= 0 else None,
                    'first_char_x': round(first_x, 2),
                    'cell_start_x': round(cell_start_x, 2),
                    'left_indent': round(left_indent, 2),
                    'right_indent': round(right_indent, 2),
                    'first_line_indent': round(first_line_indent, 2),
                    'char_spacing': round(cs, 4),
                    'chars': len(text),
                    'text': text[:30],
                }
                results.append(entry)

                available = cell_w - left_indent - right_indent
                needed = len(text) * 10.5  # rough: all CJK at 10.5pt
                fits = "FIT" if needed <= available else "OVERFLOW"

                print(f"R{r}C{c}: cell_w={cell_w:.1f} lp={lp:.1f} rp={rp:.1f} indent_l={left_indent:.1f} indent_r={right_indent:.1f} fi={first_line_indent:.1f} cs={cs:.2f}")
                print(f"  available={available:.1f} needed(est)={needed:.1f} {fits} \"{text[:25]}\"")

            except Exception as e:
                print(f"R{r}C{c}: error {e}")

    doc.Close(SaveChanges=False)
    word.Quit()

    out = "tools/metrics/output/cell_available_width.json"
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved to {out}")


if __name__ == "__main__":
    main()
