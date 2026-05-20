"""S138: Diagnose a1d6 +11-13pt step jumps at i=178/254/257/271/286.

For each jump-trigger paragraph:
1. Identify Word's text from COM (paragraph index → text + properties)
2. Identify Oxi's matched layout element (text + para_idx + cell)
3. Show the row/cell structure around it
4. Compare row height computed by Oxi vs Word
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', 'a1d6e4efa2e7_tokumei_08_01-4.docx')
LAYOUT = r'c:\tmp\a1d6_bundle_layout.json'

# Word para indices where jumps occur (per element_iou_diff/a1d6e4efa2e7.json bundle data)
JUMP_INDICES = [178, 254, 257, 271, 286]


def measure_word():
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(DOCX, ReadOnly=True)
    out = {'paras': []}
    try:
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            rng = p.Range
            cr = d.Range(rng.Start, rng.Start)
            txt = (rng.Text or '').strip()
            in_tbl = bool(cr.Information(12))
            row_idx = int(cr.Information(13)) if in_tbl else None  # wdWithInTable=12, RowNumber=13
            col_idx = int(cr.Information(15)) if in_tbl else None
            sb = round(p.Format.SpaceBefore, 3)
            sa = round(p.Format.SpaceAfter, 3)
            ls = round(p.Format.LineSpacing, 3)
            lsr = int(p.Format.LineSpacingRule)
            out['paras'].append({
                'i': i,
                'text': txt[:35],
                'page': int(cr.Information(3)),
                'y': round(cr.Information(6), 3),
                'x': round(cr.Information(5), 3),
                'in_table': in_tbl,
                'row': row_idx,
                'col': col_idx,
                'sb': sb,
                'sa': sa,
                'ls': ls,
                'lsr': lsr,
            })
    finally:
        d.Close(False)
        word.Quit()
    return out


def main():
    w = measure_word()
    paras = w['paras']
    # Index by i
    by_i = {p['i']: p for p in paras}

    print('=== a1d6 jump paragraphs (Word side) ===')
    for ji in JUMP_INDICES:
        # Show context: i-2, i-1, i, i+1, i+2
        print(f'\n--- jump i={ji} ---')
        for delta in range(-3, 4):
            i = ji + delta
            if i in by_i:
                p = by_i[i]
                marker = '** JUMP **' if delta == 0 else ''
                row_col = f'r{p["row"]} c{p["col"]}' if p['in_table'] else 'body'
                print(f'  i={p["i"]:>4} pg{p["page"]} y={p["y"]:>7.2f} {row_col:>7} sb={p["sb"]:>5.2f} sa={p["sa"]:>5.2f} ls={p["ls"]:>6.2f} lsr={p["lsr"]} text={p["text"]!r:>40} {marker}')

    # Save full Word measurement for cross-ref
    out_path = os.path.join(REPO, 'pipeline_data', 'a1d6_word_paras_full.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(w, f, ensure_ascii=False, indent=2)
    print(f'\nWrote {out_path}')


if __name__ == '__main__':
    main()
