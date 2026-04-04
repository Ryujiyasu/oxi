"""COM: Measure per-character widths in table cells where Oxi might overflow.

For each table cell with justify alignment, measure:
- Cell available width (cell.Width - left/right padding)
- Sum of individual character widths via horizontal position
- Whether there's slack or overflow
"""
import win32com.client
import os, time, json


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    docs = [
        "459f05f1e877_kyodokenkyuyoushiki01.docx",
        "a1d6e4efa2e7_tokumei_08_01-4.docx",
    ]

    all_results = []

    for docname in docs:
        path = os.path.abspath(f"tools/golden-test/documents/docx/{docname}")
        if not os.path.exists(path):
            continue

        print(f"\n=== {docname} ===")
        doc = word.Documents.Open(path, ReadOnly=True)
        time.sleep(1)

        for t in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(t)

            for r in range(1, tbl.Rows.Count + 1):
                for c in range(1, tbl.Columns.Count + 1):
                    try:
                        cell = tbl.Cell(r, c)
                        rng = cell.Range
                        text = rng.Text.rstrip('\r\x07')
                        if not text or len(text) < 5:
                            continue

                        para = rng.Paragraphs(1)
                        if para.Alignment != 3:  # justify only
                            continue

                        cell_w = cell.Width
                        font_name = rng.Font.Name
                        font_size = rng.Font.Size

                        # Measure first/last char horizontal positions
                        chars = rng.Characters
                        n = chars.Count

                        # Find first non-whitespace char position
                        first_x = None
                        last_x = None
                        last_w = 0

                        # Get positions of first and last text characters
                        valid_chars = []
                        for i in range(1, n + 1):
                            ch = chars(i).Text
                            if ch in ('\r', '\x07'):
                                continue
                            valid_chars.append(i)

                        if len(valid_chars) < 2:
                            continue

                        # First char left edge
                        first_ch = chars(valid_chars[0])
                        first_x = first_ch.Information(5)  # wdHorizontalPositionRelativeToPage

                        # Last char right edge
                        last_ch = chars(valid_chars[-1])
                        last_ch_end = last_ch.Duplicate
                        last_ch_end.Collapse(0)  # wdCollapseEnd
                        last_x = last_ch_end.Information(5)

                        text_span = last_x - first_x

                        # Also get cell left edge
                        cell_rng = cell.Range.Duplicate
                        cell_rng.Collapse(1)  # wdCollapseStart
                        cell_left = cell_rng.Information(5)

                        # Get individual char widths for first 10 chars
                        sample_widths = []
                        for idx in valid_chars[:10]:
                            ch_rng = chars(idx)
                            x1 = ch_rng.Information(5)
                            ch_end = ch_rng.Duplicate
                            ch_end.Collapse(0)
                            x2 = ch_end.Information(5)
                            w = x2 - x1
                            sample_widths.append({
                                'char': chars(idx).Text,
                                'width_pt': round(w, 3),
                            })

                        # Padding: cell_left - table edge, and right padding
                        # Rough available width
                        ratio = text_span / cell_w if cell_w > 0 else 0

                        entry = {
                            'doc': docname,
                            'table': t, 'row': r, 'col': c,
                            'cell_w': round(cell_w, 2),
                            'text_span': round(text_span, 2),
                            'ratio': round(ratio, 4),
                            'chars': len(valid_chars),
                            'font': font_name,
                            'font_size': font_size,
                            'text': text[:30],
                            'sample_widths': sample_widths,
                        }
                        all_results.append(entry)

                        if ratio > 0.85:
                            print(f"  DENSE T{t}R{r}C{c}: span={text_span:.1f}/{cell_w:.1f}={ratio:.3f} {len(valid_chars)}ch {font_name} {font_size}pt \"{text[:25]}\"")

                    except Exception as e:
                        pass

        doc.Close(SaveChanges=False)

    word.Quit()

    out = "tools/metrics/output/cell_char_widths.json"
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved {len(all_results)} cells to {out}")


if __name__ == "__main__":
    main()
