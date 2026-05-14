"""Measure Oxi's per-page-per-cell paragraph count for ed025c 損益計算書 table.

Companion to measure_ed025c_cell_pagination.py (Word side).

Reads the layout dump from oxi-gdi-renderer and aggregates text elements by
(page, cell_row_idx, cell_col_idx, cell_para_idx). Compares per-page-per-cell
counts vs Word's measurement to identify rows where Oxi over-fits vs Word.
"""
from __future__ import annotations
import os, sys, json, subprocess, tempfile
sys.stdout.reconfigure(encoding='utf-8')

REPO = r'c:\Users\ryuji\oxi-main'
DOC = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', 'ed025cbecffb_index-23.docx')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
WORD_DATA = os.path.join(REPO, 'pipeline_data', 'ra_manual_measurements', 'ed025c_cell_pagination.json')
OUT = os.path.join(REPO, 'pipeline_data', 'ra_manual_measurements', 'ed025c_oxi_cell_pagination.json')


def main():
    with tempfile.TemporaryDirectory(prefix='oxi_dump_') as tmp:
        out_prefix = os.path.join(tmp, 'page_')
        dump_path = os.path.join(tmp, 'layout.json')
        proc = subprocess.run(
            [RENDERER, DOC, out_prefix, '--dump-layout=' + dump_path],
            capture_output=True, text=True, timeout=120,
        )
        if proc.returncode != 0 or not os.path.exists(dump_path):
            print(f'Renderer failed: rc={proc.returncode}  {proc.stderr[:500]}')
            return
        with open(dump_path, encoding='utf-8') as f:
            dump = json.load(f)

    # Aggregate text elements by (page, cell_row_idx, cell_col_idx)
    from collections import defaultdict
    per_cell_per_page = defaultdict(lambda: defaultdict(list))
    for page in dump.get('pages', []):
        pg_num = page['page']
        # Group text elements by (cell_row_idx, cell_col_idx, cell_para_idx)
        # so multi-fragment paragraphs collapse to single paragraph records.
        grouped = defaultdict(lambda: {'y': float('inf'), 'text': ''})
        for el in page.get('elements', []):
            if el.get('type') != 'text': continue
            cri = el.get('cell_row_idx')
            cci = el.get('cell_col_idx')
            cpi = el.get('cell_para_idx')
            if cri is None or cci is None: continue
            key = (cri, cci, cpi)
            g = grouped[key]
            if el['y'] < g['y']: g['y'] = el['y']
            g['text'] += el.get('text', '')
        for (cri, cci, cpi), g in grouped.items():
            per_cell_per_page[(cri, cci)][pg_num].append({
                'cpi': cpi,
                'text': g['text'][:30],
                'y': round(g['y'], 2),
            })

    summaries = []
    for (cri, cci), pgs in sorted(per_cell_per_page.items()):
        for pg in sorted(pgs.keys()):
            paras = sorted(pgs[pg], key=lambda p: p['cpi'] if p['cpi'] is not None else 0)
            ys = [p['y'] for p in paras]
            if ys:
                first_y, last_y = min(ys), max(ys)
                span = last_y - first_y
            else:
                first_y = last_y = span = -1
            summaries.append({
                'cell_row': cri,
                'cell_col': cci,
                'page': pg,
                'n_paragraphs': len(paras),
                'first_y': first_y,
                'last_y': last_y,
                'span_pt': round(span, 2),
                'first_text': paras[0]['text'] if paras else '',
                'last_text': paras[-1]['text'] if paras else '',
            })

    # Save
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump({'doc': 'ed025cbecffb', 'summary': summaries}, f, ensure_ascii=False, indent=2)
    print(f'Saved to {OUT}')
    print()
    print('=== Oxi per-page-per-cell ===')
    print(f'{"row":>3} {"col":>3} {"page":>4} {"n":>4} {"first_y":>7} {"last_y":>7} {"span":>6}  first_text')
    for s in summaries:
        print(f'{s["cell_row"]:>3} {s["cell_col"]:>3} {s["page"]:>4} {s["n_paragraphs"]:>4} {s["first_y"]:>7} {s["last_y"]:>7} {s["span_pt"]:>6}  {s["first_text"][:30]!r}')

    # Cross-compare with Word
    if not os.path.exists(WORD_DATA):
        print(f'\nWord data not found at {WORD_DATA}; skip cross-compare')
        return
    with open(WORD_DATA, encoding='utf-8') as f:
        word = json.load(f)
    word_sum = word['per_page_per_cell_summary']
    # Build map: (word_row-1, word_col-1) → list of (page, n)
    # (Word is 1-based; Oxi is 0-based. Approximate mapping.)
    print('\n=== Word vs Oxi (paragraph count per page per cell) ===')
    # Map Word's pages (50/51/52) to Oxi's (13/14/15) — by sorted order
    word_pages = sorted(set(s['page'] for s in word_sum if s['n_paragraphs'] > 1))
    oxi_pages = sorted(set(s['page'] for s in summaries if s['cell_row'] == max(s['cell_row'] for s in summaries)))
    print(f'Word pages used by big-cell rows: {word_pages}')
    print(f'Oxi pages used by biggest row: {oxi_pages[:10]}')

    # Match by cell column (Word col 1-4 → Oxi col 0-3)
    word_by_col = defaultdict(dict)  # col_0based -> {oxi_page: word_count}
    # Use only the row with max paragraphs (the big content row in row 2)
    word_big_row = max((s['cell_row'] for s in word_sum if s['n_paragraphs'] > 1), default=None)
    if word_big_row:
        for s in word_sum:
            if s['cell_row'] == word_big_row:
                col0 = s['cell_col'] - 1
                # Map word_page to oxi-equivalent: find index in word_pages, use same index in oxi_pages
                try:
                    idx = word_pages.index(s['page'])
                    if idx < len(oxi_pages):
                        oxi_pg = oxi_pages[idx]
                        word_by_col[col0][oxi_pg] = s['n_paragraphs']
                except ValueError:
                    pass

    oxi_big_row = max((s['cell_row'] for s in summaries), default=None)
    print(f'\n{"col":>3} {"page":>4} {"Word_n":>6} {"Oxi_n":>5} {"delta":>5}')
    for col0 in sorted(word_by_col.keys()):
        for oxi_pg in sorted(word_by_col[col0].keys()):
            word_n = word_by_col[col0][oxi_pg]
            oxi_n = next((s['n_paragraphs'] for s in summaries
                          if s['cell_row'] == oxi_big_row
                          and s['cell_col'] == col0
                          and s['page'] == oxi_pg), 0)
            print(f'{col0:>3} {oxi_pg:>4} {word_n:>6} {oxi_n:>5} {oxi_n-word_n:+5d}')


if __name__ == '__main__':
    main()
