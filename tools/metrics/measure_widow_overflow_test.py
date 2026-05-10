"""Day 33 part 29 — Measure widow_overflow repros via Word COM."""
from __future__ import annotations
import os, sys, glob
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

REPRO_DIR = 'tools/golden-test/repros/widow_overflow'


def measure(docx_path):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    abs_path = os.path.abspath(docx_path)
    d = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        n = d.Paragraphs.Count
        # Find the LAST fill paragraph and the test paragraph (last in doc)
        # Test paragraph is the last non-empty paragraph
        test_idx = None
        last_fill_idx = None
        for i in range(n, 0, -1):
            p = d.Paragraphs(i)
            text = (p.Range.Text or '').strip()
            if text and not text.startswith('F'):
                test_idx = i
                break
        for i in range(n, 0, -1):
            p = d.Paragraphs(i)
            text = (p.Range.Text or '').strip()
            if text.startswith('F'):
                last_fill_idx = i
                break
        # Get positions
        last_fill_p = d.Paragraphs(last_fill_idx)
        test_p = d.Paragraphs(test_idx)
        cr_fill = d.Range(last_fill_p.Range.Start, last_fill_p.Range.Start)
        cr_test = d.Range(test_p.Range.Start, test_p.Range.Start)
        fill_pg = int(cr_fill.Information(3))
        fill_y = round(cr_fill.Information(6), 2)
        test_pg = int(cr_test.Information(3))
        test_y = round(cr_test.Information(6), 2)
        return {
            'last_fill_idx': last_fill_idx, 'last_fill_pg': fill_pg, 'last_fill_y': fill_y,
            'test_idx': test_idx, 'test_pg': test_pg, 'test_y': test_y,
            'n_pages': int(d.ComputeStatistics(2)),  # wdStatisticPages
        }
    finally:
        d.Close(False)
        word.Quit()


def main():
    repros = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
    print(f'{"variant":<40} {"last_fill":<22} {"test_para":<22} {"n_pages"}')
    for r in repros:
        name = os.path.splitext(os.path.basename(r))[0]
        try:
            m = measure(r)
        except Exception as e:
            print(f'{name}: ERROR {e}')
            continue
        fill_s = f'pg={m["last_fill_pg"]} y={m["last_fill_y"]:>6}'
        test_s = f'pg={m["test_pg"]} y={m["test_y"]:>6}'
        print(f'{name:<40} {fill_s:<22} {test_s:<22} {m["n_pages"]}')


if __name__ == '__main__':
    main()
