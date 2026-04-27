"""Measure MS Mincho NOKERN variants vs original.

Tests whether <w:kern> is the gate that triggers Word's yakumono
compression for MS Mincho. If NOKERN variants show width=10.5pt for
`、` (uncompressed) while original MC_A_mincho shows 5.25pt
(compressed), then kerning is the discriminator.
"""
import json
import os
import sys

import win32com.client


DOCS = [
    ('MC_A_mincho_ORIG_kern_compat14', 'tools/metrics/mincho_adjacency_repro/MC_A_mincho.docx'),
    ('MC_A_mincho_NOKERN_compat14', 'tools/metrics/mincho_kern_variants/MC_A_mincho_NOKERN.docx'),
    ('MC_A_mincho_NOKERN_compat15', 'tools/metrics/mincho_kern_variants/MC_A_mincho_NOKERN_COMPAT15.docx'),
]


def measure(word, docx_path):
    abs_path = os.path.abspath(docx_path)
    doc = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        para = doc.Paragraphs(1)
        r = para.Range
        text = r.Text
        per_char = []
        for i in range(len(text)):
            sub = doc.Range(r.Start + i, r.Start + i + 1)
            try:
                x = sub.Information(5)
                y = sub.Information(6)
                per_char.append({'i': i, 'ch': text[i], 'x': x, 'y': y})
            except Exception:
                per_char.append({'i': i, 'ch': text[i], 'x': None, 'y': None})
        return per_char
    finally:
        doc.Close(False)


def widths(per_char):
    """For text 観、「測 repeated, A_widths = 、 widths, B_widths = 「 widths."""
    aw, bw = [], []
    n = len(per_char)
    for cycle in range(10):
        base = cycle * 4
        if base + 2 >= n:
            break
        a_rec = per_char[base + 1]  # 、
        b_rec = per_char[base + 2]  # 「
        post_a = per_char[base + 2] if base + 2 < n else None  # for 、 width
        post_b = per_char[base + 3] if base + 3 < n else None  # for 「 width
        if a_rec['x'] is None or b_rec['x'] is None:
            continue
        if a_rec['y'] is None or b_rec['y'] is None or abs(a_rec['y'] - b_rec['y']) > 3:
            continue
        aw.append(b_rec['x'] - a_rec['x'])
        if post_b and post_b['x'] is not None and abs(post_b['y'] - b_rec['y']) < 3:
            bw.append(post_b['x'] - b_rec['x'])
    return aw, bw


def main():
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    out = {}
    try:
        for label, path in DOCS:
            print(f'measuring {label} ...', flush=True)
            per_char = measure(word, path)
            aw, bw = widths(per_char)
            out[label] = {
                'A_widths_、': aw,
                'B_widths_「': bw,
                'A_avg': sum(aw)/len(aw) if aw else None,
                'B_avg': sum(bw)/len(bw) if bw else None,
            }
    finally:
        word.Quit()
    os.makedirs('pipeline_data', exist_ok=True)
    with open('pipeline_data/mincho_kern_variants.json', 'w', encoding='utf-8') as f:
        json.dump(out, f, indent=2, ensure_ascii=False)

    print('\n== Result ==')
    print(f'{"label":<40s} {"A_avg(、)":>10s} {"B_avg(「)":>10s} {"verdict":<15s}')
    for label, r in out.items():
        a = r['A_avg']
        b = r['B_avg']
        if a is None:
            verdict = 'NO DATA'
        elif 4.5 <= a <= 6.0:
            verdict = 'COMPRESSED'
        elif 9.5 <= a <= 11.5:
            verdict = 'FULL'
        else:
            verdict = f'OTHER({a:.2f})'
        a_str = f'{a:.2f}' if a is not None else '--'
        b_str = f'{b:.2f}' if b is not None else '--'
        print(f'{label:<40s} {a_str:>10s} {b_str:>10s} {verdict:<15s}')


if __name__ == '__main__':
    sys.exit(main())
