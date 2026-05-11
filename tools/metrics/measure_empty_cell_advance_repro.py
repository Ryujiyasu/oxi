"""Day 33 part 51 — Measure empty cell paragraph advance for R7.3."""
from __future__ import annotations
import os, sys, subprocess, re
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc
import glob

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools/golden-test/repros/empty_cell_advance')
RENDERER = os.path.abspath(os.path.join(REPO, 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe'))
WD_VPOS = 6


def measure_word(docx_path):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    rows = []
    try:
        if d.Tables.Count < 1: return rows
        t = d.Tables(1)
        for r in range(1, t.Rows.Count + 1):
            cell = t.Cell(r, 1)
            rng = cell.Range
            first = d.Range(rng.Start, rng.Start)
            y = round(first.Information(WD_VPOS), 2)
            rows.append({'r': r, 'y': y})
    finally:
        d.Close(False)
        word.Quit()
    return rows


def render_oxi(docx_path):
    log = r'C:\tmp\ec_dump.log'
    env = dict(os.environ); env['OXI_DUMP_TABLE'] = '1'
    with open(log, 'w') as f:
        subprocess.run([RENDERER, os.path.abspath(docx_path), r'C:\tmp\ec',
                        '--dump-layout=' + r'C:\tmp\ec_layout.json'],
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
    print(f'{"variant":<35} {"w_r1_y":>7} {"w_r2_y":>7} {"w_adv":>7} {"o_rh":>7} {"diff":>6}')
    for docx in variants:
        name = os.path.splitext(os.path.basename(docx))[0]
        try:
            word_rows = measure_word(docx)
            oxi_rows = render_oxi(docx)
        except Exception as e:
            print(f'{name}: ERR {e}')
            continue
        if len(word_rows) < 2 or len(oxi_rows) < 2: continue
        w_adv = word_rows[1]['y'] - word_rows[0]['y']
        o_rh = oxi_rows[0]['rh']
        diff = o_rh - w_adv
        print(f'{name:<35} {word_rows[0]["y"]:>7} {word_rows[1]["y"]:>7} '
              f'{w_adv:>+7.2f} {o_rh:>+7.2f} {diff:>+6.2f}')


if __name__ == '__main__':
    main()
