"""COM: Measure exact character positions in justify table cells.

For cells where Word uses justify, measure every character's X position
to determine the actual inter-character spacing Word applies.
Compare justify vs left alignment.
"""
import win32com.client
import os, time, json


def measure_char_positions(rng, max_chars=50):
    """Get X position of each character."""
    chars = rng.Characters
    n = min(chars.Count, max_chars)
    positions = []
    for i in range(1, n + 1):
        ch = chars(i).Text
        if ch in ('\r', '\x07'):
            continue
        x = chars(i).Information(5)  # wdHorizontalPositionRelativeToPage
        positions.append({'char': ch, 'x': round(x, 2), 'idx': i})
    return positions


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    path = os.path.abspath("tools/golden-test/documents/docx/459f05f1e877_kyodokenkyuyoushiki01.docx")
    doc = word.Documents.Open(path, ReadOnly=False)
    time.sleep(1)

    results = []

    # Test specific cells: T1R2C2 (13 chars, ratio=1.001)
    test_cells = [
        (1, 2, 2),  # 13ch, ratio=1.001
        (1, 3, 2),  # 13ch, ratio=1.001
        (1, 4, 2),  # 15ch, ratio=1.001
        (1, 5, 2),  # 8ch, ratio=1.001
        (1, 8, 2),  # 42ch, ratio=1.127 (overflow!)
    ]

    for t, r, c in test_cells:
        try:
            cell = doc.Tables(t).Cell(r, c)
            rng = cell.Range
            text = rng.Text.rstrip('\r\x07')
            para = rng.Paragraphs(1)
            cell_w = cell.Width

            # Measure with justify
            justify_pos = measure_char_positions(rng)

            # Switch to left
            para.Alignment = 0  # left
            time.sleep(0.1)
            left_pos = measure_char_positions(rng)

            # Restore
            para.Alignment = 3  # justify
            time.sleep(0.1)

            # Calculate spacing
            if len(justify_pos) >= 2:
                j_gaps = []
                for i in range(1, len(justify_pos)):
                    gap = justify_pos[i]['x'] - justify_pos[i-1]['x']
                    j_gaps.append(round(gap, 3))

                l_gaps = []
                for i in range(1, len(left_pos)):
                    gap = left_pos[i]['x'] - left_pos[i-1]['x']
                    l_gaps.append(round(gap, 3))

                # First char indent
                j_indent = justify_pos[0]['x']
                l_indent = left_pos[0]['x']

                entry = {
                    'table': t, 'row': r, 'col': c,
                    'cell_w': round(cell_w, 2),
                    'text': text[:30],
                    'chars': len(text),
                    'justify_gaps': j_gaps,
                    'left_gaps': l_gaps,
                    'justify_start_x': j_indent,
                    'left_start_x': l_indent,
                    'justify_avg_gap': round(sum(j_gaps)/len(j_gaps), 3) if j_gaps else 0,
                    'left_avg_gap': round(sum(l_gaps)/len(l_gaps), 3) if l_gaps else 0,
                }
                results.append(entry)

                print(f"\nT{t}R{r}C{c}: \"{text[:25]}\" ({len(text)}ch, cell_w={cell_w:.1f})")
                print(f"  Justify: start={j_indent:.1f} avg_gap={entry['justify_avg_gap']:.3f}")
                print(f"    gaps: {j_gaps[:15]}")
                print(f"  Left:   start={l_indent:.1f} avg_gap={entry['left_avg_gap']:.3f}")
                print(f"    gaps: {l_gaps[:15]}")

                # Natural text width (left-aligned)
                if left_pos:
                    last_ch_rng = rng.Characters(len(left_pos))
                    last_end = last_ch_rng.Duplicate
                    last_end.Collapse(0)
                    natural_w = last_end.Information(5) - left_pos[0]['x']
                    print(f"  Natural text width: {natural_w:.1f}pt")
                    print(f"  Cell available: {cell_w:.1f} - 2*5.76 = {cell_w - 11.52:.1f}pt")
                    print(f"  Overflow: {natural_w - (cell_w - 11.52):.2f}pt")

        except Exception as e:
            print(f"T{t}R{r}C{c}: error {e}")

    doc.Close(SaveChanges=False)
    word.Quit()

    out = "tools/metrics/output/justify_spacing_data.json"
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved to {out}")


if __name__ == "__main__":
    main()
