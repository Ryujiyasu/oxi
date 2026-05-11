"""Day 33 part 36 — Measure factor C minimal repros: per-row Word vs Oxi y.

For each variant, output a per-row trajectory:
  row N: word_y, oxi_y, dx, delta_per_row

Tests hypothesis (-2.35pt/row for v_align=center + lh=exact + fs=10.5)
and isolates which attribute drives it.
"""
from __future__ import annotations
import os, sys, json, subprocess, re
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'factor_c')
RENDERER = os.path.abspath(os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe'))


VARIANTS = [
    'v01_center_exact240_fs10p5',
    'v02_top_exact240_fs10p5',
    'v03_center_auto_fs10p5',
    'v04_center_exact240_fs12',
    'v05_center_exact300_fs10p5',
]


def render_oxi(docx):
    label = os.path.splitext(os.path.basename(docx))[0]
    out_layout = os.path.join(r'C:\tmp', f'factor_c_{label}_layout.json')
    cmd = [RENDERER, docx, os.path.join(r'C:\tmp', f'factor_c_{label}'),
           f'--dump-layout={out_layout}']
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if r.returncode != 0:
        print(f'    Oxi render returncode={r.returncode}')
        print(f'    stderr={r.stderr[-500:]}')
    with open(out_layout, encoding='utf-8') as f:
        return json.load(f)


def measure_oxi(layout):
    """Return list of (label, y) for text elements like 'row01', 'row02'..."""
    rows = []
    for page in layout.get('pages', []):
        pg = page.get('page')
        for el in page.get('elements', []):
            if el.get('type') != 'text': continue
            text = (el.get('text') or '').strip()
            m = re.match(r'^row(\d+)$', text)
            if m:
                rows.append({
                    'r': int(m.group(1)),
                    'pg': pg,
                    'y': round(el.get('y', 0), 2),
                    'x': round(el.get('x', 0), 2),
                })
    rows.sort(key=lambda r: r['r'])
    return rows


def measure_word(docx):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
    rows = []
    try:
        if d.Tables.Count < 1:
            print('    (no tables in doc)')
            return rows
        t = d.Tables(1)
        n_rows = t.Rows.Count
        for r in range(1, n_rows + 1):
            try:
                cell = t.Cell(r, 1)
                rng = cell.Range
                first = d.Range(rng.Start, rng.Start)
                y = round(first.Information(6), 2)
                pg = int(first.Information(3))
                text = (rng.Text or '').replace('\r', ' ').replace('\x07', '').strip()
            except Exception as e:
                y, pg, text = -1, -1, f'err: {e}'
            try: top_pad = round(float(cell.TopPadding), 2)
            except: top_pad = -1
            try: v_align = int(cell.VerticalAlignment)
            except: v_align = -1
            try: cell_h = round(float(cell.Height), 2)
            except: cell_h = -1
            try: row_h = round(float(t.Rows(r).Height), 2)
            except: row_h = -1
            try: row_h_rule = int(t.Rows(r).HeightRule)
            except: row_h_rule = -1
            rows.append({
                'r': r,
                'pg': pg,
                'y': y,
                'text': text[:20],
                'top_pad': top_pad,
                'v_align': v_align,
                'cell_h': cell_h,
                'row_h': row_h,
                'row_h_rule': row_h_rule,
            })
    finally:
        d.Close(False)
        word.Quit()
    return rows


def process(variant):
    docx = os.path.join(REPRO_DIR, f'{variant}.docx')
    if not os.path.exists(docx):
        print(f'  {variant}: NOT FOUND')
        return
    print(f'\n=== {variant} ===')
    layout = render_oxi(docx)
    oxi_rows = measure_oxi(layout)
    word_rows = measure_word(docx)

    # Cross-join by r
    oxi_by_r = {r['r']: r for r in oxi_rows}
    print(f'  {"r":>3} {"word_y":>8} {"oxi_y":>8} {"dx":>7}  {"delta":>6}  row_h_rule={word_rows[0]["row_h_rule"] if word_rows else "?"}')
    prev_dx = None
    deltas = []
    for w in word_rows:
        ox = oxi_by_r.get(w['r'])
        if not ox:
            print(f'  r{w["r"]:>2}  word_y={w["y"]:>6}  (no Oxi match)')
            continue
        dx = round(ox['y'] - w['y'], 2)
        delta_str = '       '
        if prev_dx is not None:
            delta = round(dx - prev_dx, 2)
            delta_str = f'{delta:+6.2f}'
            deltas.append(delta)
        print(f'  r{w["r"]:>2} {w["y"]:>8} {ox["y"]:>8}  {dx:+7.2f}  {delta_str}')
        prev_dx = dx
    if deltas:
        mean_d = sum(deltas) / len(deltas)
        print(f'  per-row delta: n={len(deltas)}  mean={mean_d:+.3f}pt  '
              f'min={min(deltas):+.2f}  max={max(deltas):+.2f}')


def main():
    print('Factor C minimal repro measurement campaign')
    print(f'Renderer: {RENDERER}')
    print(f'Repro dir: {REPRO_DIR}')
    if not os.path.exists(RENDERER):
        print('ERROR: Oxi GDI renderer binary not found')
        return
    targets = sys.argv[1:] if len(sys.argv) > 1 else VARIANTS
    for v in targets:
        try:
            process(v)
        except Exception as e:
            print(f'  {v}: ERROR {e}')


if __name__ == '__main__':
    main()
