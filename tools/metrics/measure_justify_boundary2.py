"""COM: Check alignment of all table cells in problem docs."""
import win32com.client
import os, time, json
from collections import Counter


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    docs = [
        "459f05f1e877_kyodokenkyuyoushiki01.docx",
        "a1d6e4efa2e7_tokumei_08_01-4.docx",
    ]

    for docname in docs:
        path = os.path.abspath(f"tools/golden-test/documents/docx/{docname}")
        if not os.path.exists(path):
            continue

        print(f"\n=== {docname} ===")
        doc = word.Documents.Open(path, ReadOnly=True)
        time.sleep(1)

        align_counts = Counter()
        multi_line_cells = []

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
                        if not text:
                            continue

                        # Check all paragraphs in cell
                        for p in range(1, rng.Paragraphs.Count + 1):
                            para = rng.Paragraphs(p)
                            align_counts[para.Alignment] += 1

                        # Count lines
                        lines = rng.ComputeStatistics(1)
                        cell_w = cell.Width

                        if lines >= 2 and len(text) > 3:
                            multi_line_cells.append({
                                'table': t, 'row': r, 'col': c,
                                'cell_w': round(cell_w, 2),
                                'lines': lines,
                                'chars': len(text),
                                'align': rng.Paragraphs(1).Alignment,
                                'text': text[:40],
                            })

                        if lines == 1 and len(text) > 10:
                            # Interesting: many chars on 1 line
                            ratio = len(text) * 6.0 / cell_w  # rough estimate
                            if ratio > 0.8:
                                multi_line_cells.append({
                                    'table': t, 'row': r, 'col': c,
                                    'cell_w': round(cell_w, 2),
                                    'lines': lines,
                                    'chars': len(text),
                                    'align': rng.Paragraphs(1).Alignment,
                                    'text': text[:40],
                                    'note': 'dense_1line'
                                })

                    except Exception as e:
                        pass

        align_names = {0: 'left', 1: 'center', 2: 'right', 3: 'justify', 4: 'distribute'}
        print(f"Alignment distribution:")
        for k, v in align_counts.most_common():
            print(f"  {align_names.get(k, k)}: {v}")

        print(f"\nMulti-line or dense cells ({len(multi_line_cells)}):")
        for mc in multi_line_cells[:20]:
            note = mc.get('note', '')
            print(f"  T{mc['table']}R{mc['row']}C{mc['col']}: {mc['lines']}lines {mc['chars']}ch w={mc['cell_w']}pt align={mc['align']} {note} \"{mc['text'][:30]}\"")

        doc.Close(SaveChanges=False)

    word.Quit()


if __name__ == "__main__":
    main()
