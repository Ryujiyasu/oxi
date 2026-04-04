"""COM: Measure exact justify compression behavior.

For each table cell with justify alignment:
1. Switch to LEFT alignment
2. Check if first/last char Y positions differ (= 2+ lines)
3. Switch back to JUSTIFY
4. Check if all chars are on same Y (= 1 line, compressed)
5. If left=multiline and justify=1line, measure compression ratio

This uses Y position differences to detect line breaks,
avoiding unreliable ComputeStatistics.
"""
import win32com.client
import os, time, json


def get_line_info(rng, max_chars=200):
    """Get Y positions of characters to determine line structure."""
    chars = rng.Characters
    n = min(chars.Count, max_chars)

    lines = []  # list of {'y': float, 'chars': int, 'start_x': float, 'end_x': float}
    current_y = None
    current_count = 0
    current_start_x = None
    current_end_x = None

    for i in range(1, n + 1):
        ch = chars(i).Text
        if ch in ('\r', '\x07', '\n'):
            continue
        x = chars(i).Information(5)  # wdHorizontalPositionRelativeToPage
        y = round(chars(i).Information(6), 1)  # wdVerticalPositionRelativeToPage

        if current_y is None or abs(y - current_y) > 1.0:
            if current_y is not None:
                lines.append({
                    'y': current_y,
                    'chars': current_count,
                    'start_x': current_start_x,
                    'end_x': current_end_x,
                })
            current_y = y
            current_count = 0
            current_start_x = x

        current_count += 1
        current_end_x = x

    if current_y is not None and current_count > 0:
        lines.append({
            'y': current_y,
            'chars': current_count,
            'start_x': current_start_x,
            'end_x': current_end_x,
        })

    return lines


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    docs = [
        "459f05f1e877_kyodokenkyuyoushiki01.docx",
        "a1d6e4efa2e7_tokumei_08_01-4.docx",
    ]

    results = []

    for docname in docs:
        path = os.path.abspath(f"tools/golden-test/documents/docx/{docname}")
        if not os.path.exists(path):
            continue

        print(f"\n=== {docname} ===")
        doc = word.Documents.Open(path, ReadOnly=False)
        time.sleep(1)

        for t in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(t)
            for r in range(1, tbl.Rows.Count + 1):
                for c in range(1, tbl.Columns.Count + 1):
                    try:
                        cell = tbl.Cell(r, c)
                        rng = cell.Range
                        text = rng.Text.rstrip('\r\x07')
                        if not text or len(text) < 3:
                            continue

                        para = rng.Paragraphs(1)
                        orig_align = para.Alignment
                        if orig_align != 3:  # justify only
                            continue

                        cell_w = cell.Width
                        font_size = rng.Font.Size

                        # Measure with JUSTIFY
                        justify_lines = get_line_info(rng)

                        # Switch to LEFT
                        para.Alignment = 0
                        time.sleep(0.05)
                        left_lines = get_line_info(rng)

                        # Restore
                        para.Alignment = 3
                        time.sleep(0.05)

                        j_line_count = len(justify_lines)
                        l_line_count = len(left_lines)

                        # Compression detected: left wraps but justify doesn't
                        compressed = (l_line_count > j_line_count)

                        if compressed:
                            # Natural text width = end_x of last char on last line (left aligned)
                            # For more accuracy, sum all lines
                            natural_w = 0
                            for ll in left_lines:
                                line_w = ll['end_x'] - ll['start_x']
                                natural_w += line_w
                            # Approximate: sum of first-line width + overflow
                            first_line_w = left_lines[0]['end_x'] - left_lines[0]['start_x']
                            overflow_chars = sum(ll['chars'] for ll in left_lines[1:])

                            entry = {
                                'doc': docname,
                                'table': t, 'row': r, 'col': c,
                                'cell_w': round(cell_w, 2),
                                'total_chars': len(text),
                                'justify_lines': j_line_count,
                                'left_lines': l_line_count,
                                'first_line_chars_left': left_lines[0]['chars'],
                                'overflow_chars': overflow_chars,
                                'font_size': font_size,
                                'text': text[:40],
                            }
                            results.append(entry)

                            print(f"  COMPRESS T{t}R{r}C{c}: j_lines={j_line_count} l_lines={l_line_count} "
                                  f"cell_w={cell_w:.1f} total_ch={len(text)} "
                                  f"left_l1_ch={left_lines[0]['chars']} overflow={overflow_chars} "
                                  f"\"{text[:30]}\"")

                    except Exception as e:
                        pass

        doc.Close(SaveChanges=False)

    word.Quit()

    # Summary
    print(f"\n=== Summary ===")
    print(f"Total compressed cells: {len(results)}")
    if results:
        ratios = []
        for r in results:
            # Rough overflow ratio: total_chars / first_line_chars_left
            if r['first_line_chars_left'] > 0:
                ratio = r['total_chars'] / r['first_line_chars_left']
                ratios.append(ratio)
                print(f"  {r['doc']} T{r['table']}R{r['row']}C{r['col']}: "
                      f"{r['total_chars']}ch/{r['first_line_chars_left']}ch_left "
                      f"= {ratio:.3f} cell_w={r['cell_w']}")

    out = "tools/metrics/output/justify_compress_exact.json"
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"Saved to {out}")


if __name__ == "__main__":
    main()
