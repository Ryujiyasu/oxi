"""Measure every line Y position on db9c page 2 for drift analysis."""

import win32com.client
import os
import json
import time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc_path = os.path.abspath("tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx")
    doc = word.Documents.Open(doc_path, ReadOnly=True)
    time.sleep(1)

    # Find all paragraphs on page 2
    lines = []
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        rng = p.Range
        start_char = doc.Range(rng.Start, rng.Start + 1)
        page = start_char.Information(3)

        if page < 2:
            continue
        if page > 2:
            # Check if paragraph spans from page 2
            end_char = doc.Range(rng.End - 2, rng.End - 1)
            end_page = end_char.Information(3)
            if end_page > 2:
                # Get lines on page 2 only
                pass
            else:
                break

        # Track each line in this paragraph
        y = start_char.Information(6)
        text = rng.Text[:60].replace('\r', '').replace('\n', '')

        # Find line breaks by character Y changes
        prev_y = None
        line_start = rng.Start
        for ci in range(rng.Start, min(rng.End, rng.Start + 800)):
            cr = doc.Range(ci, ci + 1)
            try:
                cy = cr.Information(6)
                cp = cr.Information(3)
            except:
                continue

            if cp > 2:
                break

            if prev_y is None or abs(cy - prev_y) > 1.0:
                if prev_y is not None:
                    lines.append({
                        'para': i,
                        'y': prev_y,
                        'chars': ci - line_start,
                    })
                line_start = ci
                prev_y = cy

        # Last line
        if prev_y is not None:
            lines.append({
                'para': i,
                'y': prev_y,
                'chars': min(rng.End, rng.Start + 800) - line_start,
            })

    doc.Close(False)
    word.Quit()

    print(f"Lines on page 2: {len(lines)}")
    for l in lines:
        print(f"  P{l['para']} y={l['y']:.1f} [{l['chars']}c]")

    # Save
    with open('pipeline_data/db9c_page2_lines.json', 'w') as f:
        json.dump(lines, f, indent=2)

if __name__ == '__main__':
    measure()
