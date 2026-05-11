"""Day 33 part 44 — Per-cell line-count Word vs Oxi for de6e.

Hypothesis (from Day 33 part 42-43): the +3105pt cell-content overhang
is from multi-line wrapped text. If Oxi's line counts per cell differ
from Word's, that's the dominant compaction site.

For each cell in de6e (all tables, gridSpan-safe via try/except):
- Word: cell.Range.ComputeStatistics(wdStatisticLines) — authoritative
  line count
- Word: cell.Height (rendered cell box height)
- Word: first paragraph fs, lh, sb, sa, lineRule
- Oxi: count text elements falling within cell border bounds in layout JSON

Output: pipeline_data/de6e_cell_line_counts.csv

Cells where Word_lines != Oxi_lines = divergent wrap.
"""
from __future__ import annotations
import os, sys, json, subprocess, csv
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = os.path.join(REPO, 'tools/golden-test/documents/docx/de6e32b5960b_tokumei_08_01-1.docx')
RENDERER = os.path.join(REPO, 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')

WD_STAT_LINES = 1  # wdStatisticLines


def render_oxi(docx, force=False):
    label = os.path.splitext(os.path.basename(docx))[0]
    out_layout = os.path.join(r'C:\tmp', f'de6e_clc_{label}_layout.json')
    if force or not os.path.exists(out_layout):
        cmd = [RENDERER, docx, os.path.join(r'C:\tmp', f'de6e_clc_{label}'),
               f'--dump-layout={out_layout}']
        subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    with open(out_layout, encoding='utf-8') as f:
        return json.load(f)


def measure_word(docx_path):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    cells = []
    try:
        n_tables = d.Tables.Count
        for t_idx in range(1, n_tables + 1):
            try: t = d.Tables(t_idx)
            except Exception: continue
            try: n_rows = t.Rows.Count
            except: continue
            try: n_cols = t.Columns.Count
            except: continue
            for r in range(1, n_rows + 1):
                for c in range(1, n_cols + 1):
                    try: cell = t.Cell(r, c)
                    except Exception: continue
                    try:
                        rng = cell.Range
                        first = d.Range(rng.Start, rng.Start)
                        y = round(first.Information(6), 2)
                        pg = int(first.Information(3))
                    except Exception:
                        y, pg = -1, -1
                    try: cell_h = round(float(cell.Height), 2)
                    except: cell_h = -1
                    # ComputeStatistics(wdStatisticLines) returns 0 for cell ranges
                    # (Word's quirk). Use cell_h / lh as proxy.
                    try:
                        n_lines = int(rng.ComputeStatistics(WD_STAT_LINES))
                    except Exception:
                        n_lines = -1
                    try: v_align = int(cell.VerticalAlignment)
                    except: v_align = -1
                    try: text = (rng.Text or '').replace('\r', ' ').replace('\x07', '').strip()[:40]
                    except: text = ''
                    try: p1 = rng.Paragraphs(1)
                    except: p1 = None
                    fs, lh, sb, sa, lh_rule = -1, -1, -1, -1, -1
                    if p1 is not None:
                        try: fs = float(p1.Range.Font.Size)
                        except: pass
                        try: lh = round(float(p1.Format.LineSpacing), 2)
                        except: pass
                        try: lh_rule = int(p1.Format.LineSpacingRule)
                        except: pass
                        try: sb = round(float(p1.Format.SpaceBefore), 2)
                        except: pass
                        try: sa = round(float(p1.Format.SpaceAfter), 2)
                        except: pass
                    # Proxy: cell_h / lh ≈ word's rendered line count
                    if cell_h > 0 and lh > 0:
                        word_n_lines_proxy = round((cell_h - (sb if sb > 0 else 0) - (sa if sa > 0 else 0)) / lh)
                        word_n_lines_proxy = max(1, word_n_lines_proxy)
                    else:
                        word_n_lines_proxy = -1
                    cells.append({
                        't': t_idx, 'r': r, 'c': c, 'pg': pg, 'y': y,
                        'word_n_lines': n_lines, 'word_n_lines_proxy': word_n_lines_proxy,
                        'cell_h': cell_h,
                        'v_align': v_align, 'fs': fs, 'lh': lh, 'sb': sb, 'sa': sa,
                        'lh_rule': lh_rule, 'text': text,
                    })
    finally:
        d.Close(False)
        word.Quit()
    return cells


def count_oxi_lines_per_cell(layout, word_cells):
    """For each Word cell, count Oxi text elements roughly inside cell area.

    Cell bounding box: (cell_top, cell_left, cell_top+cell_h, cell_left+cell_w).
    We approximate cell_top via word_cell.y - top_pad (typically 0-0.5pt).
    Since we don't know cell_left/right precisely from Word COM here,
    use looser y-range filter: text elements within y±cell_h centered on word y.
    """
    # Group Oxi text elements by page
    by_page = {}
    for page in layout.get('pages', []):
        pg = page.get('page')
        texts = []
        for el in page.get('elements', []):
            if el.get('type') == 'text':
                texts.append({
                    'pg': pg,
                    'y': round(el.get('y', 0), 2),
                    'x': round(el.get('x', 0), 2),
                    'text': el.get('text', '') or '',
                    'para_idx': el.get('para_idx'),
                })
        by_page[pg] = texts

    for wc in word_cells:
        if wc['pg'] < 1 or wc['y'] < 0:
            wc['oxi_n_lines'] = -1
            continue
        # Look at Oxi text within y±cell_h/2 of Word y on same page
        page_texts = by_page.get(wc['pg'], [])
        y_lo = wc['y'] - 1.0
        y_hi = wc['y'] + max(wc['cell_h'], 12.0) + 1.0
        in_range = [t for t in page_texts if y_lo <= t['y'] <= y_hi]
        # Count unique y-lines (rounded to 1pt)
        unique_lines = set(round(t['y']) for t in in_range)
        wc['oxi_n_lines'] = len(unique_lines)
        wc['oxi_text_sample'] = '|'.join(t['text'][:8] for t in in_range[:5])


def main():
    print('Measuring de6e per-cell line counts...')
    word_cells = measure_word(DOCX)
    print(f'  Word: {len(word_cells)} accessible cells')
    layout = render_oxi(DOCX, force=True)
    count_oxi_lines_per_cell(layout, word_cells)

    out_path = os.path.join(REPO, 'pipeline_data', 'de6e_cell_line_counts.csv')
    with open(out_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=list(word_cells[0].keys()))
        writer.writeheader()
        writer.writerows(word_cells)
    print(f'Wrote {out_path}')

    # Summary: cells where Word_lines_proxy != Oxi_lines
    divergent = [c for c in word_cells if c['word_n_lines_proxy'] > 0 and c['oxi_n_lines'] >= 0 and c['word_n_lines_proxy'] != c['oxi_n_lines']]
    print(f'\nDivergent cells (Word_proxy != Oxi): {len(divergent)} of {len(word_cells)}')
    if divergent:
        divergent.sort(key=lambda c: -abs(c['word_n_lines_proxy'] - c['oxi_n_lines']))
        print(f'\nTop 20 divergent cells:')
        print(f'  t,r,c   pg  word_p oxi_l diff cell_h fs lh    text')
        for c in divergent[:20]:
            diff = c['oxi_n_lines'] - c['word_n_lines_proxy']
            print(f'  t{c["t"]}r{c["r"]:>2}c{c["c"]:>2}  {c["pg"]:>2}  '
                  f'{c["word_n_lines_proxy"]:>5} {c["oxi_n_lines"]:>5} {diff:+5} '
                  f'{c["cell_h"]:>6} {c["fs"]:>4} {c["lh"]:>5}  '
                  f'{c["text"][:30]!r}')

        from collections import defaultdict
        by_lh = defaultdict(list)
        for c in word_cells:
            if c['word_n_lines_proxy'] > 0 and c['oxi_n_lines'] >= 0:
                by_lh[c['lh']].append(c['oxi_n_lines'] - c['word_n_lines_proxy'])
        print('\nBy lh:')
        for lh, diffs in sorted(by_lh.items()):
            if len(diffs) >= 3:
                avg = sum(diffs) / len(diffs)
                n_pos = sum(1 for d in diffs if d > 0)
                n_neg = sum(1 for d in diffs if d < 0)
                n_zero = sum(1 for d in diffs if d == 0)
                print(f'  lh={lh}: n={len(diffs)} mean_diff={avg:+.2f}  same={n_zero} oxi_more={n_pos} oxi_less={n_neg}')

    total_proxy = sum(c['word_n_lines_proxy'] for c in word_cells if c['word_n_lines_proxy'] > 0)
    total_oxi = sum(c['oxi_n_lines'] for c in word_cells if c['oxi_n_lines'] >= 0)
    print(f'\nTotal Word lines proxy: {total_proxy}')
    print(f'Total Oxi lines: {total_oxi}')
    print(f'Diff: {total_oxi - total_proxy:+} lines (each line ~12pt -> {(total_oxi - total_proxy)*12} pt of over-pump)')


if __name__ == '__main__':
    main()
