"""Day 33 part 52 — Per-row Word vs Oxi advance for 备考 cluster (R7.4).

Generalizes measure_de6e_per_row_full.py to run on all 6 备考 cluster docs.
Captures per-row data, filters out cross-page artifacts (|diff|>=30),
aggregates per-row diff distribution across docs.

Output: pipeline_data/cluster_per_row_<doc>.csv + summary.
"""
from __future__ import annotations
import os, sys, csv, subprocess, re
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc
import glob
from collections import Counter, defaultdict

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
RENDERER = os.path.abspath(os.path.join(REPO, 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe'))
PAGE_H = 841.95
WD_VPOS = 6
WD_PAGE = 3

CLUSTER = ['de6e32b5960b', 'd4d126dfe1d9', '6514f214e482',
           'a1d6e4efa2e7', '191cb5254cb2', '1636d28e2c46']


def abs_y(pg, y):
    if pg is None or pg < 1 or y is None or y < 0:
        return None
    return (pg - 1) * PAGE_H + y


def measure_word_per_row(docx_path):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    rows_data = []
    try:
        for t_idx in range(1, d.Tables.Count + 1):
            t = d.Tables(t_idx)
            try: n_rows = t.Rows.Count; n_cols = t.Columns.Count
            except: continue
            for r in range(1, n_rows + 1):
                cell_data = None
                for c in range(1, n_cols + 1):
                    try:
                        cell = t.Cell(r, c)
                        rng = cell.Range
                        first = d.Range(rng.Start, rng.Start)
                        y = round(first.Information(WD_VPOS), 2)
                        pg = int(first.Information(WD_PAGE))
                        try: row_h = round(float(t.Rows(r).Height), 2)
                        except: row_h = -1
                        try: row_h_rule = int(t.Rows(r).HeightRule)
                        except: row_h_rule = -1
                        try: v_align = int(cell.VerticalAlignment)
                        except: v_align = -1
                        try: p1 = rng.Paragraphs(1)
                        except: p1 = None
                        fs = lh = sb = sa = lh_rule = -1
                        if p1 is not None:
                            try: fs = float(p1.Range.Font.Size)
                            except: pass
                            try: lh = round(float(p1.Format.LineSpacing), 2)
                            except: pass
                            try: sb = round(float(p1.Format.SpaceBefore), 2)
                            except: pass
                            try: sa = round(float(p1.Format.SpaceAfter), 2)
                            except: pass
                            try: lh_rule = int(p1.Format.LineSpacingRule)
                            except: pass
                        cell_data = {'t': t_idx, 'r': r, 'pg': pg, 'y': y,
                                     'row_h_rule': row_h_rule, 'v_align': v_align,
                                     'fs': fs, 'lh': lh, 'sb': sb, 'sa': sa, 'lh_rule': lh_rule}
                        break
                    except Exception:
                        continue
                if cell_data:
                    rows_data.append(cell_data)
    finally:
        d.Close(False)
        word.Quit()
    return rows_data


def parse_oxi_dump(docx_path):
    log_path = r'C:\tmp\cluster_dump.log'
    cmd = [RENDERER, os.path.abspath(docx_path), r'C:\tmp\cluster_out',
           '--dump-layout=' + r'C:\tmp\cluster_layout.json']
    env = dict(os.environ); env['OXI_DUMP_TABLE'] = '1'
    with open(log_path, 'w') as f:
        subprocess.run(cmd, stderr=f, stdout=subprocess.DEVNULL, env=env, timeout=120)
    tables = []
    current = None
    with open(log_path, encoding='utf-8') as f:
        for line in f:
            m = re.match(r'\[TBL_DUMP\] row=(\d+) entry_cursor_y=([\d.]+) row_height_pre=([\d.]+) trHeight=([\d.]+) rule=(\S+) n_cells=(\d+)', line)
            if m:
                row, cy, rh, trh, rule, ncells = int(m.group(1)), float(m.group(2)), float(m.group(3)), float(m.group(4)), m.group(5), int(m.group(6))
                if row == 0:
                    if current: tables.append(current)
                    current = []
                current.append({'row': row, 'cy': cy, 'rh': rh, 'trh': trh, 'rule': rule})
    if current: tables.append(current)
    return tables


def measure_doc(doc_id):
    docx_path = glob.glob(f'tools/golden-test/documents/docx/{doc_id}*')
    if not docx_path:
        print(f'  {doc_id}: NOT FOUND')
        return None
    docx_path = docx_path[0]
    print(f'  Measuring {doc_id} ({os.path.basename(docx_path)})...')
    word_rows = measure_word_per_row(docx_path)
    oxi_tables = parse_oxi_dump(docx_path)
    # Build (t, r) → oxi row
    oxi_map = {}
    for t_idx, rows in enumerate(oxi_tables, 1):
        for r in rows:
            oxi_map[(t_idx, r['row'] + 1)] = r

    # Compute Word per-row advance + diff
    out_rows = []
    for i, wr in enumerate(word_rows):
        word_ay = abs_y(wr['pg'], wr['y'])
        next_word_ay = None
        if i + 1 < len(word_rows) and word_rows[i+1]['t'] == wr['t']:
            next_word_ay = abs_y(word_rows[i+1]['pg'], word_rows[i+1]['y'])
        word_adv = (next_word_ay - word_ay) if (word_ay and next_word_ay) else None
        oxi_r = oxi_map.get((wr['t'], wr['r']))
        oxi_rh = oxi_r['rh'] if oxi_r else None
        diff = (oxi_rh - word_adv) if (oxi_rh and word_adv) else None
        out_rows.append({
            **wr,
            'doc': doc_id,
            'word_adv': round(word_adv, 2) if word_adv else '',
            'oxi_rh': round(oxi_rh, 2) if oxi_rh else '',
            'diff': round(diff, 2) if diff else '',
        })

    # Save
    out_path = f'pipeline_data/cluster_per_row_{doc_id}.csv'
    with open(out_path, 'w', encoding='utf-8', newline='') as f:
        if out_rows:
            writer = csv.DictWriter(f, fieldnames=list(out_rows[0].keys()))
            writer.writeheader()
            writer.writerows(out_rows)
    return out_rows


def main():
    print('Cluster per-row measurement (R7.4):')
    all_rows = []
    for doc_id in CLUSTER:
        rows = measure_doc(doc_id)
        if rows: all_rows.extend(rows)

    # Aggregate: filter clean (|diff|<30)
    clean = [r for r in all_rows if r['diff'] != '' and abs(float(r['diff'])) < 30]
    big = [r for r in all_rows if r['diff'] != '' and abs(float(r['diff'])) >= 30]
    print(f'\n=== Aggregate ({len(clean)} clean + {len(big)} cross-page rows) ===')

    # Bucket clean by lh_rule × v_align × row_h_rule
    print(f'\nClean per-row diff by (lh_rule, row_h_rule, v_align):')
    buckets = defaultdict(list)
    for r in clean:
        key = (r['lh_rule'], r['row_h_rule'], r['v_align'])
        buckets[key].append(float(r['diff']))
    for key, diffs in sorted(buckets.items(), key=lambda kv: -len(kv[1])):
        if len(diffs) < 3: continue
        mean = sum(diffs) / len(diffs)
        lh_r, rh_r, va = key
        lh_name = {0: 'single', 1: '1.5x', 2: 'dbl', 3: 'atleast', 4: 'exact', 5: 'mul'}.get(lh_r, '?')
        rh_name = {0: 'auto', 1: 'atLeast', 2: 'exact'}.get(rh_r, '?')
        va_name = {0: 'top', 1: 'center', 3: 'bottom'}.get(va, '?')
        print(f'  lh={lh_name:>8} row={rh_name:>8} v={va_name:>6} n={len(diffs):>3}  mean={mean:+5.2f}pt  min={min(diffs):+5.2f}  max={max(diffs):+5.2f}')

    # Per-doc summary
    print(f'\nPer-doc summary:')
    for doc_id in CLUSTER:
        doc_rows = [r for r in clean if r['doc'] == doc_id]
        if not doc_rows: continue
        sum_diff = sum(float(r['diff']) for r in doc_rows)
        print(f'  {doc_id}: {len(doc_rows)} clean rows, sum diff = {sum_diff:+.1f}pt')


if __name__ == '__main__':
    main()
