"""Day 33 part 30 — Measure db9ca_replica variants."""
from __future__ import annotations
import os, sys, glob
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

REPRO_DIR = 'tools/golden-test/repros/db9ca_replica'


def measure(docx_path):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    abs_path = os.path.abspath(docx_path)
    d = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        n = d.Paragraphs.Count
        # Test paragraph is the last in doc
        test_p = d.Paragraphs(n)
        test_r = test_p.Range
        # Get first character's y
        first_c = d.Range(test_r.Start, test_r.Start)
        test_pg = int(first_c.Information(3))
        test_y = round(first_c.Information(6), 2)
        # Last fill paragraph
        fill_idx = n - 1
        fill_p = d.Paragraphs(fill_idx)
        fill_r = fill_p.Range
        cr_fill = d.Range(fill_r.Start, fill_r.Start)
        fill_pg = int(cr_fill.Information(3))
        fill_y = round(cr_fill.Information(6), 2)
        # n pages
        n_pg = int(d.ComputeStatistics(2))
        return {
            'last_fill_pg': fill_pg, 'last_fill_y': fill_y,
            'test_pg': test_pg, 'test_y': test_y, 'n_pages': n_pg,
        }
    finally:
        d.Close(False)
        word.Quit()


def main():
    repros = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
    print(f'{"variant":<42} {"last_fill":<22} {"test_para":<22} {"n_pg"}')
    for r in repros:
        name = os.path.splitext(os.path.basename(r))[0]
        try:
            m = measure(r)
        except Exception as e:
            print(f'{name}: ERROR {e}')
            continue
        fill_s = f'pg={m["last_fill_pg"]} y={m["last_fill_y"]:>6}'
        test_s = f'pg={m["test_pg"]} y={m["test_y"]:>6}'
        print(f'{name:<42} {fill_s:<22} {test_s:<22} {m["n_pages"]}')


if __name__ == '__main__':
    main()
