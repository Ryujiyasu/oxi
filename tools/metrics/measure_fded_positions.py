"""Measure Word paragraph Y positions for fded6867fcbc document."""

import win32com.client
import os
import json
import time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc_path = os.path.abspath("tools/golden-test/documents/docx/fded6867fcbc_index-15.docx")
    doc = word.Documents.Open(doc_path, ReadOnly=True)
    time.sleep(1)

    results = {}
    results['page_count'] = doc.ComputeStatistics(2)

    sec = doc.Sections(1)
    ps = sec.PageSetup
    results['page_setup'] = {
        'page_width_pt': ps.PageWidth,
        'left_margin_pt': ps.LeftMargin,
        'right_margin_pt': ps.RightMargin,
        'top_margin_pt': ps.TopMargin,
        'bottom_margin_pt': ps.BottomMargin,
        'content_width_pt': ps.PageWidth - ps.LeftMargin - ps.RightMargin,
    }

    para_data = []
    for i in range(1, min(doc.Paragraphs.Count + 1, 41)):
        p = doc.Paragraphs(i)
        rng = p.Range
        try:
            y = rng.Information(6)
            x = rng.Information(5)
            page = rng.Information(3)
        except:
            y, x, page = -1, -1, -1
        text = rng.Text[:40].replace('\r', '\\r').replace('\n', '\\n')
        para_data.append({
            'index': i,
            'y_pt': y,
            'x_pt': x,
            'page': page,
            'char_count': len(rng.Text) - 1,
            'text_preview': text,
        })

    results['paragraphs'] = para_data

    doc.Close(False)
    word.Quit()

    out_path = 'pipeline_data/fded_positions.json'
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    print(f"Pages: {results['page_count']}")
    print(f"Content width: {results['page_setup']['content_width_pt']:.2f}pt")
    print(f"Top margin: {results['page_setup']['top_margin_pt']:.2f}pt")
    print()
    for p in para_data:
        print(f"  P{p['index']}: p{p['page']} y={p['y_pt']:.1f} x={p['x_pt']:.1f} [{p['char_count']}c] {p['text_preview'][:50]}")

if __name__ == '__main__':
    measure()
