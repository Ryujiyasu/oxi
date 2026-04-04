"""COM measurement: Justify compression boundary test.

For each justify cell that has 1 line, temporarily change to Left alignment
and check if it becomes 2 lines. This tells us whether Word is compressing
the text to fit on 1 line.

If Left=2 lines and Justify=1 line, Word is applying compression.
We then measure the compression ratio.
"""
import win32com.client
import os, time, json


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    docs = [
        "459f05f1e877_kyodokenkyuyoushiki01.docx",
        "a1d6e4efa2e7_tokumei_08_01-4.docx",
        "6514f214e482_tokumei_08_01-2.docx",
        "3a4f9fbe1a83_001620506.docx",
    ]

    results = []

    for docname in docs:
        path = os.path.abspath(f"tools/golden-test/documents/docx/{docname}")
        if not os.path.exists(path):
            continue

        print(f"\n=== {docname} ===")
        # Open as read-write to allow temp alignment change
        doc = word.Documents.Open(path, ReadOnly=False)
        time.sleep(1)

        for t in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(t)
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
                        orig_align = para.Alignment

                        # Only test justify-aligned paragraphs
                        if orig_align != 3:
                            continue

                        cell_w = cell.Width

                        # Count lines with justify
                        justify_lines = rng.ComputeStatistics(1)  # wdStatisticLines

                        # Temporarily switch to left
                        para.Alignment = 0  # wdAlignParagraphLeft
                        time.sleep(0.05)
                        left_lines = rng.ComputeStatistics(1)

                        # Restore
                        para.Alignment = 3
                        time.sleep(0.05)

                        if justify_lines >= 1 and left_lines >= 1:
                            compressed = (left_lines > justify_lines)
                            entry = {
                                'doc': docname,
                                'table': t,
                                'row': r,
                                'col': c,
                                'cell_width': round(cell_w, 2),
                                'chars': len(text),
                                'justify_lines': justify_lines,
                                'left_lines': left_lines,
                                'compressed': compressed,
                                'text': text[:40],
                            }
                            results.append(entry)

                            if compressed:
                                print(f"  COMPRESSED T{t}R{r}C{c}: justify={justify_lines} left={left_lines} cell_w={cell_w:.1f}pt chars={len(text)} \"{text[:30]}\"")

                    except Exception as e:
                        pass

        # Close WITHOUT saving
        doc.Close(SaveChanges=False)

    word.Quit()

    # Summary
    compressed_cells = [r for r in results if r['compressed']]
    same_cells = [r for r in results if not r['compressed'] and r['justify_lines'] == r['left_lines']]
    print(f"\n=== Summary ===")
    print(f"Total justify cells tested: {len(results)}")
    print(f"Compressed (left>justify lines): {len(compressed_cells)}")
    print(f"Same line count: {len(same_cells)}")

    # Save
    out = "tools/metrics/output/justify_boundary_data.json"
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"Saved to {out}")


if __name__ == "__main__":
    main()
