"""Measure Word line break positions for db9ca18368cd document.
Compare Word vs Oxi line wrap points to find divergence cause."""

import win32com.client
import os
import json
import time

def measure_line_breaks():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc_path = os.path.abspath("tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx")
    doc = word.Documents.Open(doc_path, ReadOnly=True)

    time.sleep(1)

    results = {}

    # Basic doc info
    results['page_count'] = doc.ComputeStatistics(2)  # wdStatisticPages

    # Measure paragraph Y positions
    para_data = []
    for i in range(1, min(doc.Paragraphs.Count + 1, 57)):
        p = doc.Paragraphs(i)
        rng = p.Range

        try:
            y = rng.Information(6)  # wdVerticalPositionRelativeToPage
            x = rng.Information(5)  # wdHorizontalPositionRelativeToPage
            page = rng.Information(3)  # wdActiveEndPageNumber
        except:
            y, x, page = -1, -1, -1

        text = rng.Text[:80].replace('\r', '\\r').replace('\n', '\\n')

        para_data.append({
            'index': i,
            'y_pt': y,
            'x_pt': x,
            'page': page,
            'text_preview': text,
            'char_count': len(rng.Text) - 1,  # exclude paragraph mark
        })

    results['paragraphs'] = para_data

    # For key paragraphs, measure line-by-line positions
    # P7 (index 8 in 1-based) is the first long paragraph
    key_paras = [8, 10, 11, 12]  # 1-based indices for long paragraphs

    line_data = {}
    for pi in key_paras:
        if pi > doc.Paragraphs.Count:
            continue
        p = doc.Paragraphs(pi)
        rng = p.Range

        # Use character-by-character position tracking to find line breaks
        lines = []
        para_start = rng.Start
        para_end = rng.End

        prev_y = None
        line_start_offset = 0
        line_text = ""
        line_idx = 0

        for ci in range(para_start, min(para_end, para_start + 500)):
            try:
                cr = doc.Range(ci, ci + 1)
                cy = cr.Information(6)
                ch = cr.Text
            except:
                continue

            if prev_y is not None and abs(cy - prev_y) > 1.0:
                # New line detected
                lines.append({
                    'line': line_idx,
                    'start_offset': line_start_offset,
                    'y_pt': prev_y,
                    'char_count': len(line_text),
                    'text': line_text[:80],
                })
                line_idx += 1
                line_start_offset = ci - para_start
                line_text = ch if ch else ""
            else:
                line_text += ch if ch else ""

            prev_y = cy

        # Last line
        if line_text:
            lines.append({
                'line': line_idx,
                'start_offset': line_start_offset,
                'y_pt': prev_y if prev_y else 0,
                'char_count': len(line_text),
                'text': line_text[:80],
            })

        line_data[f'para_{pi}'] = lines

    results['line_breaks'] = line_data

    # Also measure content width
    # Section page setup
    sec = doc.Sections(1)
    ps = sec.PageSetup
    results['page_setup'] = {
        'page_width_pt': ps.PageWidth,
        'page_height_pt': ps.PageHeight,
        'left_margin_pt': ps.LeftMargin,
        'right_margin_pt': ps.RightMargin,
        'top_margin_pt': ps.TopMargin,
        'bottom_margin_pt': ps.BottomMargin,
        'content_width_pt': ps.PageWidth - ps.LeftMargin - ps.RightMargin,
    }

    doc.Close(False)
    word.Quit()

    # Save results
    out_path = 'pipeline_data/db9c_linebreaks.json'
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    print(f"Saved to {out_path}")
    print(f"Pages: {results['page_count']}")
    print(f"Content width: {results['page_setup']['content_width_pt']:.2f}pt")
    print(f"\nParagraph Y positions:")
    for p in para_data[:15]:
        print(f"  P{p['index']}: y={p['y_pt']:.1f} x={p['x_pt']:.1f} p{p['page']} [{p['char_count']}c] {p['text_preview'][:50]}")

    print(f"\nLine breaks:")
    for pk, lines in line_data.items():
        print(f"\n{pk} ({len(lines)} lines):")
        for l in lines[:5]:
            print(f"  L{l['line']}: off={l['start_offset']} y={l['y_pt']:.1f} [{l['char_count']}c] {l['text'][:60]}")
        if len(lines) > 5:
            print(f"  ... ({len(lines) - 5} more lines)")

if __name__ == '__main__':
    measure_line_breaks()
