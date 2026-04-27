"""Measure all 5 inverse-strip variants of 0e7af.

For each variant, find an existing yakumono pair (`、` or `。` followed
by another closing yakumono char) in the body and report its width.
The variant where width drops from ~9pt (full at 9pt body font) to
~4.5pt (compressed at 9pt) IS the one that removed the discriminator.
"""
import os
import sys

import win32com.client


VARIANTS = [
    ('V0_unmodified', 'tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx'),
    ('V1_strip_lang', 'tools/metrics/inverse_strip_variants/0e7af_V1_strip_lang.docx'),
    ('V2_strip_rPrDefault', 'tools/metrics/inverse_strip_variants/0e7af_V2_strip_rPrDefault.docx'),
    ('V3_strip_pPrDefault', 'tools/metrics/inverse_strip_variants/0e7af_V3_strip_pPrDefault.docx'),
    ('V4_strip_docDefaults', 'tools/metrics/inverse_strip_variants/0e7af_V4_strip_docDefaults.docx'),
    ('V5_minimal_styles', 'tools/metrics/inverse_strip_variants/0e7af_V5_minimal_styles.docx'),
]

CLOSING = set('、。」）．，')  # fullwidth yakumono only — halfwidth has different layout


def measure_first_yakumono(word, docx_path):
    """Find the first VALID yakumono pair in body. Skips:
    - cross-line pairs (different y)
    - cross-paragraph pairs (very wide width)
    - half-width chars (different layout rule)
    Returns (prev_ch, next_ch, prev_width, y) for first valid pair.
    """
    doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        skipped = 0
        for p_idx in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(p_idx)
            r = para.Range
            text = r.Text
            n = len(text)
            if n < 4:  # Need real paragraph, not a heading-ish blob
                continue
            for i in range(n - 1):
                if text[i] in CLOSING:
                    sub_a = doc.Range(r.Start + i, r.Start + i + 1)
                    sub_b = doc.Range(r.Start + i + 1, r.Start + i + 2)
                    try:
                        ax = sub_a.Information(5); ay = sub_a.Information(6)
                        bx = sub_b.Information(5); by = sub_b.Information(6)
                    except Exception:
                        continue
                    if ax is None or bx is None or ay is None or by is None:
                        continue
                    if abs(ay - by) > 3:
                        continue
                    width = bx - ax
                    # Skip clearly-anomalous wide widths (cross-element /
                    # wrap-induced). Valid widths for a 9-10pt font yakumono
                    # are 4-12pt.
                    if width < 0 or width > 20:
                        skipped += 1
                        continue
                    return text[i], text[i+1], width, ay, p_idx
        return None, None, None, None, None
    finally:
        doc.Close(False)


def main():
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    results = []
    try:
        for label, path in VARIANTS:
            print(f'measuring {label} ...', flush=True)
            try:
                a, b, w, y, pidx = measure_first_yakumono(word, path)
                results.append((label, a, b, w, y, pidx))
            except Exception as e:
                print(f'  ERROR: {e}', flush=True)
                results.append((label, None, None, None, None, None))
    finally:
        word.Quit()

    print('\n== Result ==')
    print(f'{"variant":<28s} {"pair":<6s} {"width":>8s} {"y":>6s} {"para":>5s} {"verdict":<25s}')
    for label, a, b, w, y, pidx in results:
        if w is None:
            print(f'{label:<28s} {"--":<6s} {"--":>8s} {"--":>6s} {"--":>5s} NO PAIR FOUND')
            continue
        pair = (a + b) if a and b else ''
        # 0e7af body is 9pt MS Mincho. Full = 9.0, Compressed = 4.5.
        if 4.0 <= w <= 5.5:
            verdict = 'COMPRESSED (gate removed!)'
        elif 8.5 <= w <= 11.5:
            verdict = 'FULL (gate still active)'
        elif 5.5 < w < 8.5:
            verdict = f'PARTIAL ({w:.2f}pt)'
        else:
            verdict = f'OTHER ({w:.2f}pt)'
        print(f'{label:<28s} {pair:<6s} {w:>8.2f} {y:>6.1f} {pidx:>5d} {verdict:<25s}')


if __name__ == '__main__':
    sys.exit(main())
