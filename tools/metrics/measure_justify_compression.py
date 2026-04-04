"""COM measurement: Word justify compression in table cells.

Goal: Determine exactly when Word compresses character spacing to fit
text on one line vs breaking to two lines in table cells.

Measures:
- Cell width
- Text content and character count
- Line count (via line-by-line enumeration)
- Character spacing (computed by Word)
- Total text width vs cell width ratio

Uses kyodokenkyuyoushiki01 and tokumei_08_01 series which have
narrow table cells with dense text.
"""
import win32com.client
import os, sys, time, json

def measure_cell_lines(cell):
    """Count lines in a cell by moving selection line by line."""
    rng = cell.Range
    # Remove end-of-cell marker
    text = rng.Text.rstrip('\r\x07')
    if not text:
        return {'lines': 1, 'text': '', 'chars': 0}

    # Count lines by using Range.ComputeStatistics
    # wdStatisticLines = 1
    try:
        lines = rng.ComputeStatistics(1)
    except:
        lines = 1

    return {
        'lines': max(lines, 1),
        'text': text[:50],  # truncate for display
        'chars': len(text),
    }


def measure_char_spacing(rng):
    """Get character spacing from range."""
    try:
        return rng.Font.Spacing  # in points
    except:
        return 0.0


def measure_table_cells(doc, table_idx):
    """Measure all cells in a table."""
    results = []
    tbl = doc.Tables(table_idx)
    rows = tbl.Rows.Count
    cols = tbl.Columns.Count

    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            try:
                cell = tbl.Cell(r, c)
                rng = cell.Range

                # Cell width
                cell_w = cell.Width

                # Cell content
                info = measure_cell_lines(cell)

                # Character spacing
                cs = measure_char_spacing(rng)

                # Font info
                font_name = rng.Font.Name
                font_size = rng.Font.Size

                # Paragraph justification
                para = rng.Paragraphs(1)
                alignment = para.Alignment  # 0=left,1=center,2=right,3=justify

                results.append({
                    'row': r,
                    'col': c,
                    'cell_width_pt': round(cell_w, 2),
                    'lines': info['lines'],
                    'chars': info['chars'],
                    'text': info['text'],
                    'char_spacing': round(cs, 4),
                    'font': font_name,
                    'font_size': font_size,
                    'alignment': alignment,
                })
            except Exception as e:
                pass

    return results


def measure_justify_details(doc, table_idx):
    """For cells with justify alignment and 1 line, measure per-character widths
    to determine compression ratio."""
    results = []
    tbl = doc.Tables(table_idx)
    rows = tbl.Rows.Count
    cols = tbl.Columns.Count

    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            try:
                cell = tbl.Cell(r, c)
                rng = cell.Range
                text = rng.Text.rstrip('\r\x07')

                if not text or len(text) < 3:
                    continue

                para = rng.Paragraphs(1)
                if para.Alignment != 3:  # justify only
                    continue

                info = measure_cell_lines(cell)
                cell_w = cell.Width

                # Measure total text width by setting alignment to left
                # and checking the line count
                # Instead: measure individual character advances
                # Using Range.Characters
                chars_rng = rng.Characters
                char_count = chars_rng.Count

                total_width = 0.0
                char_widths = []
                for i in range(1, min(char_count + 1, 200)):
                    try:
                        ch_rng = chars_rng(i)
                        ch = ch_rng.Text
                        if ch in ('\r', '\x07'):
                            continue
                        # MoveEnd to just this character, get its width via
                        # horizontal position difference
                        # wdHorizontalPositionRelativeToPage = 5
                        x1 = ch_rng.Information(5)

                        # Duplicate range and collapse to end
                        ch_rng2 = ch_rng.Duplicate
                        ch_rng2.Collapse(0)  # wdCollapseEnd
                        x2 = ch_rng2.Information(5)

                        w = x2 - x1
                        if w > 0:
                            total_width += w
                            char_widths.append(round(w, 2))
                    except:
                        pass

                if total_width > 0:
                    ratio = total_width / cell_w
                    results.append({
                        'row': r,
                        'col': c,
                        'cell_width': round(cell_w, 2),
                        'text_width': round(total_width, 2),
                        'ratio': round(ratio, 4),
                        'lines': info['lines'],
                        'chars': info['chars'],
                        'text': text[:30],
                        'char_spacing': round(rng.Font.Spacing, 4),
                    })
            except Exception as e:
                pass

    return results


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    docs_to_test = [
        "459f05f1e877_kyodokenkyuyoushiki01.docx",
        "a1d6e4efa2e7_tokumei_08_01-4.docx",
        "6514f214e482_tokumei_08_01-2.docx",
    ]

    all_results = {}

    for docname in docs_to_test:
        path = os.path.abspath(f"tools/golden-test/documents/docx/{docname}")
        if not os.path.exists(path):
            print(f"SKIP: {docname} not found")
            continue

        print(f"\n=== {docname} ===")
        doc = word.Documents.Open(path, ReadOnly=True)
        time.sleep(1)

        print(f"Tables: {doc.Tables.Count}")

        doc_results = []
        for t in range(1, doc.Tables.Count + 1):
            print(f"\n--- Table {t} ---")
            cells = measure_table_cells(doc, t)

            # Show cells where lines=1 with justify (potential compression)
            justify_1line = [c for c in cells if c['alignment'] == 3 and c['lines'] == 1 and c['chars'] > 5]
            justify_2line = [c for c in cells if c['alignment'] == 3 and c['lines'] >= 2 and c['chars'] > 5]

            if justify_1line:
                print(f"  Justify 1-line cells ({len(justify_1line)}):")
                for c in justify_1line[:5]:
                    print(f"    R{c['row']}C{c['col']}: {c['chars']}ch in {c['cell_width_pt']}pt, cs={c['char_spacing']}, \"{c['text'][:30]}\"")

            if justify_2line:
                print(f"  Justify 2+ line cells ({len(justify_2line)}):")
                for c in justify_2line[:5]:
                    print(f"    R{c['row']}C{c['col']}: {c['chars']}ch/{c['lines']}lines in {c['cell_width_pt']}pt, cs={c['char_spacing']}, \"{c['text'][:30]}\"")

            # Detailed width measurement for justify cells
            details = measure_justify_details(doc, t)
            if details:
                print(f"  Width ratios:")
                for d in details:
                    marker = "***" if d['lines'] == 1 and d['ratio'] < 1.0 else ""
                    print(f"    R{d['row']}C{d['col']}: text_w={d['text_width']}pt / cell_w={d['cell_width']}pt = {d['ratio']:.4f} lines={d['lines']} cs={d['char_spacing']} {marker}")

            doc_results.append({
                'table': t,
                'cells': cells,
                'justify_details': details,
            })

        all_results[docname] = doc_results
        doc.Close(SaveChanges=False)

    word.Quit()

    # Save raw data
    out_path = "tools/metrics/output/justify_compression_data.json"
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved to {out_path}")


if __name__ == "__main__":
    main()
