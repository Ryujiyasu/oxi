# -*- coding: utf-8 -*-
"""
Re-measure 1636 items 1-5 per-character X positions on CURRENT code state
(post Session 54 cs-snap fix). Compares Word vs Oxi to localize residual
Bug B contributors.

Method:
1. Word side: COM `Information(WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE)` per
   character within paragraphs 1-5 (the "1./2./3./4./5." numbered items at
   the top of 1636 page 1).
2. Oxi side: layout-json export of the same paragraphs and read each fragment's
   x_start. Difference per char = Bug B residual.

Output: pipeline_data/bug_b_1636_residual.json
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOC = os.path.abspath('tools/golden-test/documents/docx/1636d28e2c46_tokumei_08_04.docx')
OUT = os.path.abspath('pipeline_data/bug_b_1636_residual.json')

# Items 1-5 are the first 5 numbered paragraphs after the title (in main text).
# We need to find them dynamically — they are top-of-doc body paras with 'item'-like pattern.
# Per Day 10 memo: items 1-5 each ~50 CJK chars with similar structure.

def measure_word():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(DOC, ReadOnly=True)
    out = []
    try:
        ps = d.PageSetup
        out_cfg = {
            'pgH': ps.PageHeight,
            'pgW': ps.PageWidth,
            'top': ps.TopMargin,
            'bottom': ps.BottomMargin,
            'left': ps.LeftMargin,
            'right': ps.RightMargin,
        }
        # Find first 5 body paragraphs that contain CJK and look like numbered items.
        # 1636 structure: title at top, then items. Simply pick paragraphs that have
        # text starting with "１"/"２"/"...".
        n = d.Paragraphs.Count
        target_paras = []
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            txt = (p.Range.Text or '').strip()
            if not txt:
                continue
            # match items 1-5 by leading digit
            if txt[:1] in ('１', '２', '３', '４', '５', '1', '2', '3', '4', '5'):
                target_paras.append(i)
            if len(target_paras) >= 5:
                break

        items = []
        for pi in target_paras:
            p = d.Paragraphs(pi)
            rng = p.Range
            text = rng.Text or ''
            # measure each character's HPOS via collapsed range at each offset.
            chars = []
            for off in range(min(len(text), 80)):  # cap at 80
                ch = text[off]
                if ch in ('\r', '\x07', '\n'):
                    continue
                cr = d.Range(rng.Start + off, rng.Start + off)
                x = cr.Information(5)  # wdHorizontalPositionRelativeToPage
                y = cr.Information(6)
                page = cr.Information(3)
                chars.append({
                    'off': off, 'ch': ch, 'x': x, 'y': y, 'page': page,
                })
            # Compute per-char advance from x deltas (only same-line chars)
            advances = []
            for j in range(1, len(chars)):
                if chars[j]['y'] == chars[j-1]['y']:
                    advances.append({
                        'idx': j, 'ch': chars[j-1]['ch'], 'next_ch': chars[j]['ch'],
                        'advance': chars[j]['x'] - chars[j-1]['x'],
                    })
            items.append({
                'word_para_idx': pi,
                'text_preview': text[:30].replace('\r', ''),
                'n_chars': len(chars),
                'first_x': chars[0]['x'] if chars else None,
                'first_y': chars[0]['y'] if chars else None,
                'chars': chars[:25],  # cap output
                'advances': advances[:25],
            })
        out_cfg['items'] = items
        return out_cfg
    finally:
        d.Close(SaveChanges=False)
        word.Quit()

if __name__ == '__main__':
    print('Measuring Word side...')
    word_data = measure_word()
    print(f'Found {len(word_data["items"])} items')
    for item in word_data['items']:
        print(f'  para_idx={item["word_para_idx"]}  preview={item["text_preview"]!r}  n_chars={item["n_chars"]}')
        if item['advances']:
            avgs = [a['advance'] for a in item['advances'] if a['advance'] is not None and a['advance'] > 0]
            if avgs:
                print(f'    avg advance: {sum(avgs)/len(avgs):.4f}pt over {len(avgs)} char-pairs')
                print(f'    advances: {[round(a,2) for a in avgs[:15]]}')
    json.dump(word_data, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'Saved: {OUT}')
