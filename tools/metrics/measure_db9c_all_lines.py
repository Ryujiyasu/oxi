"""Measure every line Y on ALL pages of db9c for drift comparison."""

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

    all_lines = []
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        rng = p.Range

        # First char
        start_char = doc.Range(rng.Start, rng.Start + 1)
        y0 = start_char.Information(6)
        page0 = start_char.Information(3)

        # Track lines
        prev_y = None
        line_start = rng.Start
        for ci in range(rng.Start, min(rng.End, rng.Start + 800)):
            cr = doc.Range(ci, ci + 1)
            try:
                cy = cr.Information(6)
                cp = cr.Information(3)
            except:
                continue

            if prev_y is None or abs(cy - prev_y) > 1.0 or cp != page0:
                if prev_y is not None:
                    all_lines.append({
                        'para': i,
                        'page': page0,
                        'y': prev_y,
                        'chars': ci - line_start,
                    })
                line_start = ci
                prev_y = cy
                page0 = cp

        if prev_y is not None:
            all_lines.append({
                'para': i,
                'page': page0,
                'y': prev_y,
                'chars': min(rng.End, rng.Start + 800) - line_start,
            })

    doc.Close(False)
    word.Quit()

    # Print page 1 lines
    p1_lines = [l for l in all_lines if l['page'] == 1]
    print(f"Page 1 lines: {len(p1_lines)}")
    for l in p1_lines:
        print(f"  P{l['para']} y={l['y']:.1f} [{l['chars']}c]")

    with open('pipeline_data/db9c_all_lines.json', 'w') as f:
        json.dump(all_lines, f, indent=2)

if __name__ == '__main__':
    measure()
