"""Measure Word per-line y for the 2 Bundle v8 violation paragraphs:
- c7b923e5c616 paragraph 39 ("一審の専属的合意管轄裁判所...")
- 0e7af1ae8f21 paragraphs 164, 165, 189

Goal: determine if Word puts these paragraphs at the position Bundle v8
moves them to (= fix is right direction) or at the baseline position
(= fix is wrong direction).
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))

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
            if pi > d.Paragraphs.Count:
                continue
            p = d.Paragraphs(pi)
            r = p.Range
            text = (r.Text or '').rstrip('\r\n\x07')
            cr_start = d.Range(r.Start, r.Start)
            end_pos = max(r.Start, r.End - 1)
            cr_end = d.Range(end_pos, end_pos)
            out[pi] = {
                'i': pi,
                'text': text[:50],
                'len': len(text),
                'start_y': round(cr_start.Information(WD_VPOS), 2),
                'start_page': int(cr_start.Information(WD_PAGE)),
                'end_y': round(cr_end.Information(WD_VPOS), 2),
                'end_page': int(cr_end.Information(WD_PAGE)),
            }
    finally:
        d.Close(False)
        word.Quit()
    return out


def main():
    targets = {
        'c7b923e5c616_20240705_resources_data_outline_06.docx': [37, 38, 39, 40, 41, 42, 43, 44, 45, 46],
        '0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx': [162, 163, 164, 165, 166, 167, 187, 188, 189, 190, 191],
    }
    docs_dir = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx')
    results = {}
    for fname, paras in targets.items():
        path = os.path.join(docs_dir, fname)
        if not os.path.exists(path):
            print(f'NOT FOUND: {path}')
            continue
        print(f'=== {fname} ===')
        try:
            r = measure(path, paras)
        except Exception as e:
            print(f'  ERROR: {e}')
            continue
        for pi in paras:
            if pi in r:
                p = r[pi]
                print(f'  pi={pi:>3} pg={p["start_page"]}->{p["end_page"]} y={p["start_y"]:.2f}->{p["end_y"]:.2f} len={p["len"]} | {p["text"]!r}')
        results[fname] = r
        print()
    out_path = os.path.join(REPO, 'pipeline_data', 'v8_violations_word_measurement.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'Wrote {out_path}')


if __name__ == '__main__':
    main()
