"""Measure Word's per-line y position within selected paragraphs.

For each character in a paragraph, query Information(WD_VPOS) to get
the line's y position. Group by distinct y to count lines and their
positions. Compare to Oxi's --dump-layout per-line data.

Goal: determine exactly where Word breaks paragraph 19 of db9ca
(how many lines on page 1, what y values).

Output: pipeline_data/db9ca_word_per_line.json
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx',
                    'db9ca18368cd_20241122_resource_open_data_01.docx')

WD_HPOS = 5
WD_VPOS = 6
WD_PAGE = 3


def measure(docx_path: str, target_paragraphs: list[int]) -> dict:
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    out = {}
    try:
        for pi in target_paragraphs:
            p = d.Paragraphs(pi)
            r = p.Range
            text = (r.Text or '').rstrip('\r\n\x07')
            chars_data = []
            # Step through chars, capture y per character
            n = r.End - r.Start
            # Sample every Nth character to avoid being too slow
            stride = max(1, n // 200)  # at most 200 samples per paragraph
            seen_lines = []  # list of (y, page, first_char_idx, sample_text)
            for i in range(0, n, stride):
                pos = r.Start + i
                cr = d.Range(pos, pos)
                y = round(cr.Information(WD_VPOS), 2)
                pg = int(cr.Information(WD_PAGE))
                # If new line (y differs from last), record
                if not seen_lines or abs(seen_lines[-1][0] - y) > 1.0 or seen_lines[-1][1] != pg:
                    chars_in_pos = text[i:i+30] if i < len(text) else ''
                    seen_lines.append((y, pg, i, chars_in_pos))
                chars_data.append({'i': i, 'y': y, 'pg': pg})
            out[pi] = {
                'i': pi,
                'text_len': n,
                'lines': [{'y': y, 'page': pg, 'first_char': fi, 'sample': s} for y, pg, fi, s in seen_lines],
                'n_lines': len(seen_lines),
            }
    finally:
        d.Close(False)
        word.Quit()
    return out


def main():
    # Measure paragraphs around the page break events in db9ca
    target = [11, 15, 18, 19, 20, 25, 36, 37, 38, 43]
    print(f'Measuring paragraphs {target} in db9ca...')
    result = measure(DOCX, target)
    for pi, info in sorted(result.items()):
        print(f'\n=== Para i={pi} (text_len={info["text_len"]}) ===')
        print(f'  n_lines={info["n_lines"]}')
        for line in info['lines']:
            print(f'    page={line["page"]} y={line["y"]:.2f} first_char={line["first_char"]} | {line["sample"]!r}')
    out_path = os.path.join(REPO, 'pipeline_data', 'db9ca_word_per_line.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'\nWrote {out_path}')


if __name__ == '__main__':
    main()
