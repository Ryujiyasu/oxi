"""COM: Fast justify compression detection using Selection.MoveDown.

For each justify cell, use Selection to count actual visual lines:
1. Select cell range
2. Move cursor to start of cell
3. Count MoveDown(wdLine) moves until Y changes (= line count)
4. Compare justify line count vs left-aligned line count
"""
import win32com.client
import os, time, json

# wdLine = 5, wdCell = 12
WD_LINE = 5
WD_CELL = 12


def count_visual_lines(word, cell):
    """Count visual lines in a cell using Selection.MoveDown."""
    rng = cell.Range
    rng.Select()
    sel = word.Selection

    # Move to start of cell
    sel.HomeKey(WD_LINE)  # move to start of current line

    first_y = sel.Information(6)  # wdVerticalPositionRelativeToPage
    lines = 1
    prev_y = first_y

    for _ in range(100):
        moved = sel.MoveDown(WD_LINE, 1)
        if moved == 0:
            break
        new_y = sel.Information(6)
        # Check if still in same cell
        if abs(new_y - prev_y) < 0.1:
            break  # didn't move to new line
        # Check if moved to next cell (Y jump too large or new page)
        if abs(new_y - prev_y) > 50:
            break
        lines += 1
        prev_y = new_y

    return lines


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False

    docs = [
        "459f05f1e877_kyodokenkyuyoushiki01.docx",
        "a1d6e4efa2e7_tokumei_08_01-4.docx",
        "6514f214e482_tokumei_08_01-2.docx",
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
            for r in range(1, min(tbl.Rows.Count + 1, 30)):
                for c in range(1, tbl.Columns.Count + 1):
                    try:
                        cell = tbl.Cell(r, c)
                        rng = cell.Range
                        text = rng.Text.rstrip('\r\x07')
                        if not text or len(text) < 3:
                            continue

                        para = rng.Paragraphs(1)
                        if para.Alignment != 3:  # justify only
                            continue

                        cell_w = cell.Width
                        font_size = rng.Font.Size

                        # Count lines with JUSTIFY
                        j_lines = count_visual_lines(word, cell)

                        # Switch to LEFT and count
                        para.Alignment = 0
                        time.sleep(0.02)
                        l_lines = count_visual_lines(word, cell)

                        # Restore
                        para.Alignment = 3
                        time.sleep(0.02)

                        compressed = (l_lines > j_lines)

                        if compressed:
                            # Estimate natural text width
                            est_w = len(text) * font_size if font_size < 100 else len(text) * 10.5
                            ratio = est_w / cell_w if cell_w > 0 else 0

                            entry = {
                                'doc': docname,
                                'table': t, 'row': r, 'col': c,
                                'cell_w': round(cell_w, 2),
                                'total_chars': len(text),
                                'font_size': font_size,
                                'justify_lines': j_lines,
                                'left_lines': l_lines,
                                'est_ratio': round(ratio, 3),
                                'text': text[:40],
                            }
                            results.append(entry)
                            print(f"  COMPRESS T{t}R{r}C{c}: j={j_lines} l={l_lines} "
                                  f"w={cell_w:.1f} ch={len(text)} fs={font_size} "
                                  f"est_ratio={ratio:.3f} \"{text[:25]}\"")

                    except Exception as e:
                        pass

        doc.Close(SaveChanges=False)

    word.Quit()

    print(f"\n=== Summary ===")
    print(f"Compressed cells: {len(results)}")
    if results:
        ratios = [r['est_ratio'] for r in results]
        print(f"Est overflow ratios: min={min(ratios):.3f} max={max(ratios):.3f} avg={sum(ratios)/len(ratios):.3f}")

    out = "tools/metrics/output/justify_compress_fast.json"
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"Saved to {out}")


if __name__ == "__main__":
    main()
