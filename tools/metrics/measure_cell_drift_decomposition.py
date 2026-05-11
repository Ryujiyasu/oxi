"""Day 33 part 35 (2026-05-12) — Cell drift decomposition campaign.

Goal: for each table cell in target docs, decompose Oxi vs Word vertical drift
into (A) cell first-line text_y_offset, (B) Bug 2 body centering proxy,
(C) cell row_height cumulative drift.

Per cell, capture:
- Word side (COM):
  - cell_first_char_y           = collapsed-range Information(6) on cell.Range.Start
  - cell_first_char_pg          = Information(3)
  - cell.TopPadding             (pt)
  - cell.VerticalAlignment      (wdCellAlignVerticalTop=0 / Center=1 / Bottom=3)
  - cell_first_para_fs          = first run font size
  - cell_first_para_lh_rule
  - cell_first_para_lh_val
- Oxi side (layout JSON):
  - cell_text_first_glyph_y     = first text element matching the cell-content prefix
  - cell_top_oxi                = derived from nearest horizontal border above first glyph
  - oxi_text_y_offset           = cell_text_first_glyph_y - cell_top_oxi
- Derived:
  - drift                       = oxi_glyph_y - word_first_char_y   (absolute, page-collapsed)
  - text_y_offset_drift         = oxi_text_y_offset - (word v_align "Top" → 0pt; Center → row_h/2 - fs/2)

CSV output: pipeline_data/cell_drift_decomposition_<doc_id>.csv
Aggregate findings: stdout summary per doc.

Target docs (Day 33 part 34): 1636, d4d126, 6514, a1d6, a47e6, 191cb
"""
from __future__ import annotations
import os, sys, json, subprocess, re, csv
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX_DIR = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx')
RENDERER = os.path.abspath(os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe'))
OUT_DIR = os.path.join(REPO, 'pipeline_data')
PAGE_HEIGHT = 841.95
MATCH_PREFIX_LEN = 3  # Japanese-dense cell text is often 3-4 chars (申出番号 etc.)
TARGET_DOCS = ['1636d28e2c46', 'd4d126dfe1d9', '6514f214e482', 'a1d6e4efa2e7', 'a47e6c6b2ca1', '191cb5254cb2']


def find_docx(doc_id):
    for f in os.listdir(DOCX_DIR):
        if f.startswith(doc_id) and f.endswith('.docx'):
            return os.path.join(DOCX_DIR, f)
    return None


def normalize(s):
    if not s:
        return ''
    s = s.replace('　', ' ').replace('\r', '').replace('\x07', '').strip()
    s = re.sub(r'\s+', ' ', s)
    return s


def render_oxi(docx, force=False):
    label = os.path.splitext(os.path.basename(docx))[0]
    out_layout = os.path.join(r'C:\tmp', f'{label}_celldrift_layout.json')
    if force or not os.path.exists(out_layout):
        cmd = [RENDERER, docx, os.path.join(r'C:\tmp', f'{label}_celldrift'), f'--dump-layout={out_layout}']
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if r.returncode != 0:
            print(f'    Oxi render returncode={r.returncode}')
            print(f'    stderr={r.stderr[-500:]}')
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    return layout


def collect_oxi_index(layout):
    """Per-page indices: text elements (sorted by y, x) and horizontal border y-values."""
    pages = []
    for page in layout.get('pages', []):
        pg = page.get('page')
        texts = []
        h_borders = []
        for el in page.get('elements', []):
            if el.get('type') == 'text':
                texts.append({
                    'pg': pg,
                    'x': round(el.get('x', 0), 2),
                    'y': round(el.get('y', 0), 2),
                    'w': round(el.get('w', 0), 2),
                    'h': round(el.get('h', 0), 2),
                    'text': el.get('text', '') or '',
                    'para_idx': el.get('para_idx'),
                    'font_size': el.get('font_size', 0),
                })
            elif el.get('type') == 'border' and abs(el.get('h', 0)) < 0.01:
                h_borders.append({
                    'pg': pg,
                    'x': round(el.get('x', 0), 2),
                    'y': round(el.get('y', 0), 2),
                    'w': round(el.get('w', 0), 2),
                })
        texts.sort(key=lambda e: (e['y'], e['x']))
        h_borders.sort(key=lambda e: e['y'])
        pages.append({'pg': pg, 'texts': texts, 'h_borders': h_borders})
    return pages


def measure_word_tables(docx):
    """Per cell: positional + first-char attrs."""
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
    cells = []
    try:
        tcount = d.Tables.Count
        for t_idx in range(1, tcount + 1):
            t = d.Tables(t_idx)
            try:
                n_rows = t.Rows.Count
                n_cols = t.Columns.Count
            except Exception:
                continue
            for r in range(1, n_rows + 1):
                for c in range(1, n_cols + 1):
                    try:
                        cell = t.Cell(r, c)
                    except Exception:
                        continue
                    try:
                        rng = cell.Range
                        first = d.Range(rng.Start, rng.Start)
                        y = round(first.Information(6), 2)
                        pg = int(first.Information(3))
                        text = (rng.Text or '').replace('\r', ' ').replace('\x07', '').strip()
                    except Exception:
                        y, pg, text = -1, -1, ''
                    try: top_pad = round(float(cell.TopPadding), 2)
                    except Exception: top_pad = -1
                    try: v_align = int(cell.VerticalAlignment)
                    except Exception: v_align = -1
                    try: cell_height = round(float(cell.Height), 2)
                    except Exception: cell_height = -1
                    try: cell_height_rule = int(cell.HeightRule)
                    except Exception: cell_height_rule = -1
                    # First paragraph attrs
                    try: p1 = cell.Range.Paragraphs(1)
                    except Exception: p1 = None
                    fs = -1
                    lh_rule = -1
                    lh_val = -1
                    snap = -1
                    style_name = '?'
                    if p1 is not None:
                        try: fs = float(p1.Range.Font.Size)
                        except Exception: fs = -1
                        try: lh_rule = int(p1.Format.LineSpacingRule)
                        except Exception: lh_rule = -1
                        try: lh_val = round(float(p1.Format.LineSpacing), 2)
                        except Exception: lh_val = -1
                        try: snap = int(p1.Format.SnapToGrid)
                        except Exception: snap = -1
                        try: style_name = str(p1.Style.NameLocal)
                        except Exception: style_name = '?'
                    cells.append({
                        't_idx': t_idx,
                        'r': r,
                        'c': c,
                        'word_pg': pg,
                        'word_y': y,
                        'word_top_pad': top_pad,
                        'word_v_align': v_align,
                        'word_cell_height': cell_height,
                        'word_cell_height_rule': cell_height_rule,
                        'fs': fs,
                        'lh_rule': lh_rule,
                        'lh_val': lh_val,
                        'snap': snap,
                        'style_name': style_name,
                        'text': text[:60],
                    })
    finally:
        d.Close(False)
        word.Quit()
    return cells


def match_oxi_for_cell(cell, oxi_pages, used):
    """Match an Oxi text element to this cell.

    Strategy:
      - Short cell text (<= 8 chars normalized): require EXACT equality (Oxi text
        equals cell text). This gives high-confidence matches for header labels
        like '備考', '申出番号'.
      - Long cell text (> 8 chars): match by first 8-char prefix. Used for
        nested giant cells (e.g. 1636 row 2 with full body) where the Oxi
        side splits text into many sub-elements and we only need the first.

    Restricted to the page that Word reports the cell on (preferred); if no
    match there, fall back to other pages.

    `used` is a set of (pg, y, x) keys already claimed by earlier cells.
    """
    wt = normalize(cell['text'])
    if len(wt) < MATCH_PREFIX_LEN:
        return None
    # Word cell text is split into many sub-elements on the Oxi side. Use a
    # short (4-char) prefix and accept bidirectional prefix:
    #   - Oxi text starts with cell prefix (when Oxi is a paragraph encompassing
    #     the cell start), OR
    #   - cell text starts with Oxi text AND Oxi text >= 3 chars (when Oxi
    #     element is a sub-fragment matching the cell's first 3+ chars).
    # For very short cells (<= 4 chars), use exact equality.
    exact_mode = len(wt) <= 4
    prefix = wt if exact_mode else wt[:4]
    target_pg = cell['word_pg']
    # Restrict to the page Word reports the cell on. Cross-page cells (cascade
    # artifacts) are out of scope for the cell-drift-decomposition campaign;
    # the Phase 1 pagination diff handles those.
    candidates = [p for p in oxi_pages if p['pg'] == target_pg]
    for page in candidates:
        for t in page['texts']:
            key = (t['pg'], t['y'], t['x'])
            if key in used:
                continue
            ot = normalize(t['text'])
            if exact_mode:
                matched = (ot == wt)
            else:
                # bidirectional prefix: Oxi covers cell start OR cell starts with Oxi (>=3 chars)
                matched = (ot.startswith(prefix) or
                           (len(ot) >= 3 and wt.startswith(ot)))
            if matched:
                # Derive cell_top from nearest horizontal border above the glyph
                cell_top = None
                best_above = None
                for b in page['h_borders']:
                    if b['y'] <= t['y'] - 0.5:
                        if best_above is None or b['y'] > best_above['y']:
                            best_above = b
                if best_above is not None:
                    # Ensure x-overlap (border spans cell horizontally)
                    if best_above['x'] - 0.5 <= t['x'] <= best_above['x'] + best_above['w'] + 0.5:
                        cell_top = best_above['y']
                used.add(key)
                return {
                    'oxi_pg': t['pg'],
                    'oxi_y': t['y'],
                    'oxi_x': t['x'],
                    'oxi_font_size': t['font_size'],
                    'oxi_cell_top': cell_top,
                    'oxi_text_y_offset': round(t['y'] - cell_top, 2) if cell_top is not None else None,
                    'match_mode': 'exact' if exact_mode else 'prefix8',
                }
    return None


def process(doc_id, out_dir, force_render=False):
    docx = find_docx(doc_id)
    if not docx:
        print(f'  {doc_id}: NOT FOUND')
        return None
    print(f'  {doc_id}: rendering Oxi...')
    layout = render_oxi(docx, force=force_render)
    oxi_pages = collect_oxi_index(layout)
    print(f'    Oxi pages: {len(oxi_pages)}')
    print(f'  {doc_id}: measuring Word cells via COM...')
    cells = measure_word_tables(docx)
    print(f'    Word cells: {len(cells)}')

    rows = []
    used = set()
    for cell in cells:
        m = match_oxi_for_cell(cell, oxi_pages, used)
        if m is None:
            drift_abs = None
        else:
            w_abs = (cell['word_pg'] - 1) * PAGE_HEIGHT + cell['word_y']
            o_abs = (m['oxi_pg'] - 1) * PAGE_HEIGHT + m['oxi_y']
            drift_abs = round(o_abs - w_abs, 2)
        row = {
            'doc_id': doc_id,
            't_idx': cell['t_idx'],
            'r': cell['r'],
            'c': cell['c'],
            'word_pg': cell['word_pg'],
            'word_y': cell['word_y'],
            'word_top_pad': cell['word_top_pad'],
            'word_v_align': cell['word_v_align'],
            'word_cell_height': cell['word_cell_height'],
            'word_cell_height_rule': cell['word_cell_height_rule'],
            'fs': cell['fs'],
            'lh_rule': cell['lh_rule'],
            'lh_val': cell['lh_val'],
            'snap': cell['snap'],
            'style_name': cell['style_name'],
            'oxi_pg': m['oxi_pg'] if m else '',
            'oxi_y': m['oxi_y'] if m else '',
            'oxi_x': m['oxi_x'] if m else '',
            'oxi_font_size': m['oxi_font_size'] if m else '',
            'oxi_cell_top': m['oxi_cell_top'] if m else '',
            'oxi_text_y_offset': m['oxi_text_y_offset'] if m else '',
            'match_mode': m['match_mode'] if m else '',
            'drift_abs': drift_abs if drift_abs is not None else '',
            'matched': bool(m),
            'text': cell['text'],
        }
        rows.append(row)

    out_path = os.path.join(out_dir, f'cell_drift_decomposition_{doc_id}.csv')
    with open(out_path, 'w', encoding='utf-8', newline='') as f:
        if rows:
            writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            writer.writeheader()
            writer.writerows(rows)
    n_matched = sum(1 for r in rows if r['matched'])
    drifts = [float(r['drift_abs']) for r in rows if r['matched'] and r['drift_abs'] != '']
    n_drift_lo = sum(1 for d in drifts if 0.5 <= abs(d) < 2.0)
    n_drift_mid = sum(1 for d in drifts if 2.0 <= abs(d) < 5.0)
    n_drift_hi = sum(1 for d in drifts if abs(d) >= 5.0)
    print(f'    wrote {out_path}: {len(rows)} cells, {n_matched} matched, '
          f'drift |0.5-2|={n_drift_lo} |2-5|={n_drift_mid} |>=5|={n_drift_hi}')
    return rows


def main():
    out_dir = OUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    force = '--force' in sys.argv
    targets = TARGET_DOCS
    custom = [a for a in sys.argv[1:] if not a.startswith('--')]
    if custom:
        targets = custom
    print(f'=== Cell drift decomposition campaign — {len(targets)} doc(s) ===')
    all_rows = []
    for d in targets:
        try:
            rs = process(d, out_dir, force_render=force)
            if rs:
                all_rows.extend(rs)
        except Exception as e:
            print(f'  {d}: ERROR {e}')

    # Aggregate report
    print('\n=== Aggregate ===')
    print(f'Total cells: {len(all_rows)}')
    matched = [r for r in all_rows if r['matched'] and r['drift_abs'] != '']
    print(f'Matched: {len(matched)}')
    if matched:
        drifts = [float(r['drift_abs']) for r in matched]
        print(f'  drift mean={sum(drifts)/len(drifts):+.2f}pt '
              f'min={min(drifts):+.2f} max={max(drifts):+.2f}')
        print(f'  |drift|>=5pt cells: {sum(1 for d in drifts if abs(d)>=5)}')
        print(f'  |drift|>=2pt cells: {sum(1 for d in drifts if abs(d)>=2)}')


if __name__ == '__main__':
    main()
