"""Day 33 part 15 — COM-measure y of every paragraph (incl. empty) in
d4d126 vs 664c38 first table, to find the discriminator between the two
empty-cell-paragraph regimes:
- 664c38 regime: Word renders no-sz empties at default_fs (matches Oxi)
- d4d126 regime: Word renders no-sz empties SMALLER (Oxi over-pumps)

Output: pipeline_data/empty_cell_y_<doc_id>.json
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))


def measure(docx_path: str, label: str) -> list:
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    out = []
    try:
        n = d.Paragraphs.Count
        for i in range(1, min(n, 30) + 1):
            p = d.Paragraphs(i)
            rng = p.Range
            text = (rng.Text or '').rstrip('\r\x07')
            cr_start = d.Range(rng.Start, rng.Start)
            try:
                y = cr_start.Information(6)  # wdVerticalPositionRelativeToPage
                pg = cr_start.Information(3)  # wdActiveEndPageNumber
                in_table = bool(cr_start.Information(12))  # wdWithInTable
                fs = p.Range.Font.Size  # may return Variant if mixed
            except:
                continue
            try:
                fs_f = float(fs)
            except:
                fs_f = None
            entry = {
                'i': i,
                'pg': pg,
                'y_pt': y,
                'in_table': in_table,
                'is_empty': not text.strip(),
                'text': text[:30],
                'fs': fs_f,
            }
            out.append(entry)
            print(f'  i={i:>3} pg={pg} y={y:>7.2f} inT={in_table} empty={entry["is_empty"]} fs={fs_f} text={text[:30]!r}')
    finally:
        d.Close(SaveChanges=False)
        word.Quit()
    out_path = os.path.join(REPO, 'pipeline_data', f'empty_cell_y_{label}.json')
    json.dump(out, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'\nSaved: {out_path}\n')
    return out


def main():
    docs = [
        ('d4d126dfe1d9', 'pipeline_data/golden_per_page/d4d126dfe1d9_tokumei_08_01-3_p1.docx'),
        ('664c38001b40', 'pipeline_data/golden_per_page/664c38001b40_order_12_p1.docx'),
        ('de6e32b5960b', 'pipeline_data/golden_per_page/de6e32b5960b_tokumei_08_01-1_p1.docx'),
    ]
    for label, path in docs:
        print(f'=== {label} ===')
        measure(path, label)


if __name__ == '__main__':
    main()
