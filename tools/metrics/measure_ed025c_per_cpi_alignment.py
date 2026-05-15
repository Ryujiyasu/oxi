"""Per-cell-paragraph alignment tool for ed025c 損益計算書 table.

Cross-compares Word vs Oxi at the granularity of
(cell_row_idx, cell_col_idx, cell_paragraph_index) — emits a side-by-side
report of (page, y) for each cpi across the cells of the target table.

Why this exists:
  ed025c's 1 remaining Phase 1 outlier (wi=1540 "× × ×", word_page 14 /
  oxi_page 13, delta=-1) is the last blocker for ed025c PASS. The
  per-cell empty-paragraph misalignment hypothesis ([[session58-day37-
  per-cell-empty-paragraph-misalignment]]) needs a multi-session refactor
  of mod.rs:7168-7188 (`min_overflow_text_y` re-anchor) into a
  cpi-aligned re-anchor. This tool provides the ground-truth dataset for
  that refactor: it shows, per cpi, where Word places the paragraph vs
  where Oxi does, so a future patch can be validated against ALL cells
  of the 損益計算書 table without breaking other docs.

Output: `pipeline_data/ra_manual_measurements/ed025c_per_cpi_alignment.json`
        + readable text report on stdout

IMPORTANT (2026-05-15 finding): Word's COM `table.Range.Cells.Count = 6`
(logical 2×4 outer cells), but Oxi's dump-layout reports 69 distinct
`(cell_row_idx, cell_col_idx)` keys. The 損益計算書 is a nested table
structure — Word sees the OUTER border, Oxi sees the DEEP cell hierarchy.
Cross-join by (row, col, cpi) is therefore meaningless across the two.

This tool gathers BOTH datasets verbatim (Word-side and Oxi-side cell
hierarchies are saved independently in the JSON output) and additionally
attempts a TEXT-PREFIX-BASED cross-match for paragraphs with non-trivial
text. The structural mismatch itself is a key finding for the cpi-aligned
re-anchor refactor — any future patch must work at Oxi's cell-indexing
level, not Word's logical-cell level.

Pre-existing tools (do NOT replace):
  - `measure_ed025c_cell_pagination.py` — per-page summary (buggy: uses
    Information(1)=adjusted page → section-relative numbering ≠ pagination_diff).
  - `measure_ed025c_oxi_cell_pagination.py` — Oxi side per-page summary.
  - `measure_ed025c_cell_pitch.py` — per-cell pitch (covered Word pitch
    spec hypothesis, falsified).
  This tool COMPLEMENTS those: per-cpi alignment + structural mismatch finding.

Phase 1 status: 53/55 PASS, mean 0.9842. This tool is instrumentation
only — does NOT modify oxidocs-core or change any baseline.
"""
from __future__ import annotations
import os, sys, json, subprocess, tempfile, traceback
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')

REPO = r'c:\Users\ryuji\oxi-main'
DOC = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', 'ed025cbecffb_index-23.docx')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
OUT = os.path.join(REPO, 'pipeline_data', 'ra_manual_measurements', 'ed025c_per_cpi_alignment.json')


def measure_word():
    """COM-measure Word's per-cpi (page, y) for the 損益計算書 table.

    Uses Information(3)=wdActiveEndPageNumber for absolute page count
    (matches pagination_diff's word_page semantics).
    """
    import win32com.client as wc
    word = wc.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.ScreenUpdating = False
    doc = None
    cells_data = []
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True)
        doc.Repaginate()

        # Locate the 損益計算書 table via "Ⅰ" + "営業損益" paragraph prefix
        target_para_idx = None
        for pi in range(1, doc.Paragraphs.Count + 1):
            txt = (doc.Paragraphs(pi).Range.Text or '').strip()
            if 'Ⅰ' in txt and '営業損益' in txt:
                target_para_idx = pi
                break
        if not target_para_idx:
            raise RuntimeError('Target paragraph (Ⅰ営業損益) not found')

        para = doc.Paragraphs(target_para_idx)
        tables = para.Range.Tables
        if tables.Count == 0:
            raise RuntimeError('Target paragraph not in a table')
        table = tables(1)
        print(f'Target table: rows={table.Rows.Count}, columns={table.Columns.Count}')

        # Iterate cells flat (vMerge-safe)
        n_cells = table.Range.Cells.Count
        for ci in range(1, n_cells + 1):
            try:
                cell = table.Range.Cells(ci)
                row_idx = cell.RowIndex
                col_idx = cell.ColumnIndex
                paras = cell.Range.Paragraphs
                cell_paras = []
                for pi_in_cell in range(1, paras.Count + 1):
                    p = paras(pi_in_cell)
                    txt = (p.Range.Text or '').rstrip('\r\n\x07')[:40]
                    rng = p.Range
                    # R30 fix: collapse to start to avoid Information() returning
                    # the active-end's page for multi-page paragraphs.
                    start_rng = doc.Range(rng.Start, rng.Start)
                    try:
                        page = int(start_rng.Information(3))  # wdActiveEndPageNumber=3 (absolute)
                        y = float(start_rng.Information(6))   # wdVerticalPositionRelativeToPage=6
                    except Exception:
                        page = -1
                        y = -1.0
                    cell_paras.append({
                        'cpi': pi_in_cell - 1,  # 0-based
                        'text': txt,
                        'page': page,
                        'y': round(y, 2),
                    })
                cells_data.append({
                    # Convert to Oxi 0-based indexing (Word is 1-based)
                    'cell_row_idx': row_idx - 1,
                    'cell_col_idx': col_idx - 1,
                    'n_paragraphs': len(cell_paras),
                    'paragraphs': cell_paras,
                })
            except Exception as e:
                print(f'Cell {ci}: error {e}')

        # Page setup
        ps = doc.Sections(1).PageSetup
        page_info = {
            'page_width': round(ps.PageWidth, 2),
            'page_height': round(ps.PageHeight, 2),
            'top_margin': round(ps.TopMargin, 2),
            'bottom_margin': round(ps.BottomMargin, 2),
        }
    finally:
        if doc is not None:
            doc.Close(SaveChanges=0)
        word.Quit()
    return {'cells': cells_data, 'page_info': page_info, 'target_para_idx': target_para_idx}


def measure_oxi():
    """Render via oxi-gdi-renderer --dump-layout; extract per-cpi data."""
    with tempfile.TemporaryDirectory(prefix='oxi_dump_') as tmp:
        out_prefix = os.path.join(tmp, 'page_')
        dump_path = os.path.join(tmp, 'layout.json')
        proc = subprocess.run(
            [RENDERER, DOC, out_prefix, '--dump-layout=' + dump_path],
            capture_output=True, text=True, timeout=180,
        )
        if proc.returncode != 0 or not os.path.exists(dump_path):
            raise RuntimeError(f'Renderer failed: rc={proc.returncode}, stderr={proc.stderr[:500]}')
        with open(dump_path, encoding='utf-8') as f:
            dump = json.load(f)

    # Aggregate per (cell_row_idx, cell_col_idx, cpi): take min y across fragments,
    # concat text, take min page.
    grouped = defaultdict(lambda: {'page': 999, 'y': float('inf'), 'text': ''})
    for page in dump.get('pages', []):
        pg_num = page['page']
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            cri = el.get('cell_row_idx')
            cci = el.get('cell_col_idx')
            cpi = el.get('cell_para_idx')
            if cri is None or cci is None:
                continue
            key = (cri, cci, cpi)
            g = grouped[key]
            if pg_num < g['page']:
                g['page'] = pg_num
                g['y'] = el['y']
                g['text'] = el.get('text', '')
            elif pg_num == g['page']:
                if el['y'] < g['y']:
                    g['y'] = el['y']
                g['text'] += el.get('text', '')

    # Filter to the 損益計算書 table — heuristic: the largest cell-count cluster
    # of any single table. For ed025c specifically, the target table is the
    # one whose cell (row, col)=(0, 0) contains "Ⅰ営業損益" or similar.
    # Since dump doesn't tag tables, we collect ALL cells; cross-join with
    # Word data will filter by overlap.
    cells = defaultdict(list)
    for (cri, cci, cpi), g in grouped.items():
        if cpi is None:
            continue
        cells[(cri, cci)].append({
            'cpi': cpi,
            'page': g['page'],
            'y': round(g['y'], 2),
            'text': g['text'][:40],
        })
    # Sort each cell's paragraphs by cpi
    result = []
    for (cri, cci), paras in sorted(cells.items()):
        result.append({
            'cell_row_idx': cri,
            'cell_col_idx': cci,
            'n_paragraphs': len(paras),
            'paragraphs': sorted(paras, key=lambda p: p['cpi']),
        })
    return {'cells': result}


def cross_join_by_text(word_data, oxi_data, min_text_len=4):
    """Cross-join by paragraph text prefix.

    Word and Oxi have different cell hierarchies (Word=6 logical cells,
    Oxi=69 deep cells), so (row,col,cpi) cross-join is meaningless.
    Instead we match by text prefix — for each Word paragraph with
    non-trivial text, find the Oxi paragraph(s) with matching text prefix.

    Returns:
      aligned: list of matched (word_para, oxi_para) records
      unmatched_word: Word paragraphs that have no Oxi match
      summary: histogram + counts
    """
    # Flatten Word side: paragraphs with non-empty text
    word_paras = []
    for cell in word_data['cells']:
        for p in cell['paragraphs']:
            t = (p['text'] or '').strip()
            if len(t) >= min_text_len:
                word_paras.append({
                    'cell_row_idx': cell['cell_row_idx'],
                    'cell_col_idx': cell['cell_col_idx'],
                    'cpi': p['cpi'],
                    'text': t,
                    'page': p['page'],
                    'y': p['y'],
                })

    # Flatten Oxi side: indexed by text prefix for fast lookup
    oxi_by_text_prefix = defaultdict(list)
    for cell in oxi_data['cells']:
        for p in cell['paragraphs']:
            t = (p['text'] or '').strip()
            if len(t) >= min_text_len:
                # Use first 10 chars as bucket key
                bucket = t[:10]
                oxi_by_text_prefix[bucket].append({
                    'cell_row_idx': cell['cell_row_idx'],
                    'cell_col_idx': cell['cell_col_idx'],
                    'cpi': p['cpi'],
                    'text': t,
                    'page': p['page'],
                    'y': p['y'],
                })

    aligned = []
    unmatched_word = []
    for w in word_paras:
        bucket = w['text'][:10]
        candidates = oxi_by_text_prefix.get(bucket, [])
        # If multiple candidates, take the one with closest page (best heuristic
        # since exact y differs across cell-indexing systems).
        if not candidates:
            unmatched_word.append(w)
            continue
        # Prefer candidate with same text (full match) and closest page
        best = None
        best_score = float('inf')
        for c in candidates:
            # text equality bonus
            score = abs(c['page'] - w['page'])
            if c['text'] != w['text']:
                score += 100  # heavy penalty for non-exact text
            if score < best_score:
                best_score = score
                best = c
        if best is None:
            unmatched_word.append(w)
            continue
        aligned.append({
            'word_cell_row_idx': w['cell_row_idx'],
            'word_cell_col_idx': w['cell_col_idx'],
            'word_cpi': w['cpi'],
            'oxi_cell_row_idx': best['cell_row_idx'],
            'oxi_cell_col_idx': best['cell_col_idx'],
            'oxi_cpi': best['cpi'],
            'word_page': w['page'],
            'word_y': w['y'],
            'oxi_page': best['page'],
            'oxi_y': best['y'],
            'page_delta': best['page'] - w['page'],
            'text': w['text'][:50],
        })

    # Histogram of page_delta
    hist = defaultdict(int)
    for a in aligned:
        hist[a['page_delta']] += 1
    summary = {
        'n_word_paragraphs': len(word_paras),
        'n_oxi_paragraphs': sum(len(v) for v in oxi_by_text_prefix.values()),
        'n_aligned': len(aligned),
        'n_unmatched_word': len(unmatched_word),
        'page_delta_hist': dict(sorted(hist.items())),
    }
    return aligned, unmatched_word, summary


def main():
    print('=== Word side: COM measurement ===')
    word_data = measure_word()
    print(f'  Cells: {len(word_data["cells"])}, target_para_idx={word_data["target_para_idx"]}')

    print('=== Oxi side: --dump-layout ===')
    oxi_data = measure_oxi()
    print(f'  Cells: {len(oxi_data["cells"])}')

    print('=== Cross-join (text-prefix based) ===')
    aligned, unmatched_word, summary = cross_join_by_text(word_data, oxi_data)
    print(f'  Word non-trivial paragraphs: {summary["n_word_paragraphs"]}')
    print(f'  Oxi non-trivial paragraphs:  {summary["n_oxi_paragraphs"]}')
    print(f'  Aligned: {summary["n_aligned"]}')
    print(f'  Unmatched (Word side): {summary["n_unmatched_word"]}')
    print(f'  page_delta_hist: {summary["page_delta_hist"]}')

    # Show outliers (page_delta != 0)
    outliers = sorted([a for a in aligned if a['page_delta'] != 0],
                      key=lambda a: (a['word_page'], a['word_y']))
    print(f'\n=== Outliers ({len(outliers)}) — page_delta != 0 ===')
    for a in outliers[:60]:
        tt = a['text'].encode('ascii','replace').decode('ascii','replace')
        print(f'  word(cell={a["word_cell_row_idx"]:>2},{a["word_cell_col_idx"]:>2} cpi={a["word_cpi"]:>3}) p={a["word_page"]:>2} y={a["word_y"]:>6.1f}  '
              f'oxi(cell={a["oxi_cell_row_idx"]:>2},{a["oxi_cell_col_idx"]:>2} cpi={a["oxi_cpi"]:>3}) p={a["oxi_page"]:>2} y={a["oxi_y"]:>6.1f}  '
              f'Δp={a["page_delta"]:+d}  [{tt[:30]}]')
    if len(outliers) > 60:
        print(f'  ... ({len(outliers)-60} more)')

    # Save
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    payload = {
        'doc': os.path.basename(DOC),
        'note': (
            'Word and Oxi have INCOMPATIBLE cell-indexing for ed025c\'s 損益計算書 — '
            'Word COM table.Range.Cells.Count = 6 (logical outer 2x4); Oxi dump-layout '
            'reports 69 distinct (cell_row_idx, cell_col_idx) keys (deep nested cells). '
            'Cross-join is by text-prefix only. See file header for context.'
        ),
        'word_data': word_data,
        'oxi_data': oxi_data,
        'aligned': aligned,
        'summary': summary,
    }
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f'\nSaved to {OUT}')


if __name__ == '__main__':
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(1)
