"""Day 33 part 50 — Measure Word per-row y for spacing collapse repros."""
from __future__ import annotations
import os, sys, subprocess, json, re
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc
import glob

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools/golden-test/repros/spacing_collapse')
RENDERER = os.path.abspath(os.path.join(REPO, 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe'))
PAGE_H = 841.95
WD_VPOS = 6


def measure_doc(docx_path):
    """Get per-row first-char y."""
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    rows = []
    try:
        if d.Tables.Count < 1:
            return rows
        t = d.Tables(1)
        n_rows = t.Rows.Count
        for r in range(1, n_rows + 1):
            cell = t.Cell(r, 1)
            rng = cell.Range
            first = d.Range(rng.Start, rng.Start)
            y = round(first.Information(WD_VPOS), 2)
            rows.append({'r': r, 'y': y})
    finally:
        d.Close(False)
        word.Quit()
    return rows


def render_oxi(docx):
    out = r'C:\tmp\sc_layout.json'
    log = r'C:\tmp\sc_dump.log'
    env = dict(os.environ); env['OXI_DUMP_TABLE'] = '1'
    with open(log, 'w') as f:
        subprocess.run([RENDERER, os.path.abspath(docx), r'C:\tmp\sc',
                        f'--dump-layout={out}'],
                       stderr=f, stdout=subprocess.DEVNULL, env=env, timeout=60)
    rows = []
    with open(log, encoding='utf-8') as f:
        for line in f:
            m = re.match(r'\[TBL_DUMP\] row=(\d+) entry_cursor_y=([\d.]+) row_height_pre=([\d.]+) trHeight=([\d.]+)', line)
            if m:
                rows.append({'r': int(m.group(1)) + 1, 'cy': float(m.group(2)),
                             'rh': float(m.group(3)), 'trh': float(m.group(4))})
    return rows


def main():
    variants = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
    print(f'{"variant":<32} {"r":>2} {"word_y":>7} {"oxi_cy":>7} {"oxi_rh":>7} {"w_adv":>7} {"o_adv":>7} {"diff":>6}')
    for docx in variants:
        name = os.path.splitext(os.path.basename(docx))[0]
        try:
            word_rows = measure_doc(docx)
            oxi_rows = render_oxi(docx)
        except Exception as e:
            print(f'{name}: ERR {e}')
            continue
        # Cross-join by r
        oxi_by_r = {r['r']: r for r in oxi_rows}
        prev_w = prev_o = None
        for w in word_rows:
            o = oxi_by_r.get(w['r'])
            if o is None: continue
            w_adv = (w['y'] - prev_w) if prev_w else None
            o_adv = (o['cy'] - prev_o) if prev_o else None
            diff = (o_adv - w_adv) if (w_adv and o_adv) else None
            w_adv_s = f'{w_adv:+6.2f}' if w_adv else '       '
            o_adv_s = f'{o_adv:+6.2f}' if o_adv else '       '
            diff_s = f'{diff:+6.2f}' if diff is not None else '      '
            print(f'{name:<32} {w["r"]:>2} {w["y"]:>7} {o["cy"]:>7.2f} {o["rh"]:>7.2f} {w_adv_s} {o_adv_s} {diff_s}')
            prev_w = w['y']
            prev_o = o['cy']
        print()


if __name__ == '__main__':
    main()
