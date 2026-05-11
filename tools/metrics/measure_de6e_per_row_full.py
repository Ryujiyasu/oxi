"""Day 33 part 49 — Per-row Word vs Oxi cursor advance for de6e (all tables).

R7.1 of multi-month page-break refactor.

For each row of each table:
- Try Cell(r, 1) first. If gridSpan error, try Cell(r, 2), ..., Cell(r, n_cols).
- Get first-char y of the cell that works
- Compute Word row advance (next row first_y - this row first_y)
- Cross-reference with Oxi OXI_DUMP_TABLE row_height_pre

Output: pipeline_data/de6e_per_row_full.csv
       per-row Word vs Oxi advance, divergence categorized by attributes.
"""
from __future__ import annotations
import os, sys, csv, subprocess, re
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc
import glob

DOCX = glob.glob('tools/golden-test/documents/docx/de6e32b5960b*')[0]
RENDERER = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
PAGE_H = 841.95
WD_VPOS = 6
WD_PAGE = 3


def abs_y(pg, y):
    if pg is None or pg < 1 or y is None or y < 0:
        return None
    return (pg - 1) * PAGE_H + y


def measure_word_per_row():
    """Walk tables, for each row try cells until one works (gridSpan-safe)."""
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    rows_data = []
    try:
        n_tables = d.Tables.Count
        for t_idx in range(1, n_tables + 1):
            t = d.Tables(t_idx)
            try: n_rows = t.Rows.Count
            except: continue
            try: n_cols = t.Columns.Count
            except: continue
            for r in range(1, n_rows + 1):
                # Try each column until one works
                cell_data = None
                for c in range(1, n_cols + 1):
                    try:
                        cell = t.Cell(r, c)
                        rng = cell.Range
                        first = d.Range(rng.Start, rng.Start)
                        y = round(first.Information(WD_VPOS), 2)
                        pg = int(first.Information(WD_PAGE))
                        text = (rng.Text or '').replace('\r', ' ').replace('\x07', '').strip()[:25]
                        try: v_align = int(cell.VerticalAlignment)
                        except: v_align = -1
                        try: cell_h = round(float(cell.Height), 2)
                        except: cell_h = -1
                        try: row_h = round(float(t.Rows(r).Height), 2)
                        except: row_h = -1
                        try: row_h_rule = int(t.Rows(r).HeightRule)
                        except: row_h_rule = -1
                        try: p1 = rng.Paragraphs(1)
                        except: p1 = None
                        fs, lh, sb, sa, lh_rule = -1, -1, -1, -1, -1
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
                        cell_data = {
                            't': t_idx, 'r': r, 'c_used': c,
                            'pg': pg, 'y': y, 'text': text,
                            'v_align': v_align, 'cell_h': cell_h,
                            'row_h_word': row_h, 'row_h_rule': row_h_rule,
                            'fs': fs, 'lh': lh, 'sb': sb, 'sa': sa, 'lh_rule': lh_rule,
                        }
                        break
                    except Exception:
                        continue
                if cell_data is None:
                    rows_data.append({'t': t_idx, 'r': r, 'c_used': -1,
                                      'pg': -1, 'y': -1, 'text': '(all cells errored)'})
                else:
                    rows_data.append(cell_data)
    finally:
        d.Close(False)
        word.Quit()
    return rows_data


def parse_oxi_dump():
    """Re-run OXI_DUMP_TABLE and parse."""
    log_path = r'C:\tmp\de6e_per_row.log'
    cmd = [RENDERER, os.path.abspath(DOCX), r'C:\tmp\de6e_per_row',
           '--dump-layout=' + r'C:\tmp\de6e_per_row_layout.json']
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


def main():
    print('Measuring Word per-row...')
    word_rows = measure_word_per_row()
    print(f'  {len(word_rows)} rows captured')
    print('Parsing Oxi dump...')
    oxi_tables = parse_oxi_dump()
    print(f'  {len(oxi_tables)} tables in Oxi dump')
    # Build Oxi per-row map: (t_idx, r_idx) → row data
    # t_idx is 1-indexed, r_idx is 0-indexed in dump
    oxi_map = {}
    for t_idx, rows in enumerate(oxi_tables, 1):
        for r in rows:
            oxi_map[(t_idx, r['row'] + 1)] = r  # convert to 1-indexed

    # Compute Word per-row advance
    out_rows = []
    # Group word_rows by table
    prev_word_ay = None
    prev_t = None
    for i, wr in enumerate(word_rows):
        word_ay = abs_y(wr['pg'], wr['y'])
        # Word advance: next row's y - this row's y, within same table
        next_word_ay = None
        if i + 1 < len(word_rows):
            nr = word_rows[i + 1]
            if nr['t'] == wr['t']:
                next_word_ay = abs_y(nr['pg'], nr['y'])
        word_adv = (next_word_ay - word_ay) if (word_ay is not None and next_word_ay is not None) else None

        oxi_r = oxi_map.get((wr['t'], wr['r']))
        oxi_rh = oxi_r['rh'] if oxi_r else None

        diff = (oxi_rh - word_adv) if (oxi_rh is not None and word_adv is not None) else None

        out_rows.append({
            **wr,
            'word_adv': round(word_adv, 2) if word_adv is not None else '',
            'oxi_rh': round(oxi_rh, 2) if oxi_rh is not None else '',
            'diff': round(diff, 2) if diff is not None else '',
        })

    # Save CSV
    out_path = 'pipeline_data/de6e_per_row_full.csv'
    with open(out_path, 'w', encoding='utf-8', newline='') as f:
        if out_rows:
            writer = csv.DictWriter(f, fieldnames=list(out_rows[0].keys()))
            writer.writeheader()
            writer.writerows(out_rows)
    print(f'Wrote {out_path}')

    # Summary
    print(f'\n{"t,r":>5} {"word_y":>7} {"word_adv":>9} {"oxi_rh":>7} {"diff":>7} v_a fs   lh    sb   sa  text')
    sum_word = 0.0
    sum_oxi = 0.0
    for r in out_rows:
        if r['word_adv'] != '' and r['oxi_rh'] != '':
            sum_word += float(r['word_adv'])
            sum_oxi += float(r['oxi_rh'])
        print(f'  t{r["t"]}r{r["r"]:>2} {r["y"]:>7} {str(r["word_adv"]):>9} {str(r["oxi_rh"]):>7} '
              f'{str(r["diff"]):>7} {r["v_align"]:>3} {r["fs"]:>4} {r["lh"]:>5} '
              f'{r["sb"]:>4} {r["sa"]:>4}  {r["text"][:25]!r}')

    print(f'\nSum Word adv (with both): {sum_word:.1f}')
    print(f'Sum Oxi rh (with both):   {sum_oxi:.1f}')
    print(f'Diff (Oxi - Word):        {sum_oxi - sum_word:+.1f}pt')


if __name__ == '__main__':
    main()
