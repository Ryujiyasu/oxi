"""Day 33 part 19 — Measure y of BEFORE/AFTER markers around all-whitespace paragraph.

If Word renders WS_142 as 1 line: AFTER marker y - BEFORE marker y ≈ 36pt
  (~18pt for BEFORE + ~18pt for the 1-line WS para = 36pt advance).
If Word wraps WS_142 to 4 lines: advance ≈ 18 + 72 = 90pt.
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOCS = [
    'WS_10', 'WS_50', 'WS_100', 'WS_142', 'WS_300', 'MIX_10_TEXT', 'MIX_50_TEXT'
]

ROOT = 'tools/golden-test/repros/fullwidth_space_wrap'


def measure(docx_path):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    abs_path = os.path.abspath(docx_path)
    d = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        n = d.Paragraphs.Count
        rows = []
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            try:
                pg = int(cr.Information(3))
                y = cr.Information(6)
            except: continue
            text = (r.Text or '').rstrip('\r\n')
            rows.append((i, pg, round(y, 2), text[:30]))
        return rows
    finally:
        d.Close(False)
        word.Quit()


def main():
    print(f'{"variant":<14} {"BEFORE_y":<12} {"WS_y":<12} {"AFTER_y":<12} {"advance":<12} {"WS_lines":<10}')
    for name in DOCS:
        path = os.path.join(ROOT, f'{name}.docx')
        try:
            rows = measure(path)
        except Exception as e:
            print(f'{name}: error {e}'); continue
        # rows: para_index, page, y, text
        # Find BEFORE / WS para / AFTER
        before = next((r for r in rows if 'BEFORE' in r[3]), None)
        after = next((r for r in rows if 'AFTER' in r[3]), None)
        ws = next((r for r in rows if r is not before and r is not after), None)
        b_y = before[2] if before else None
        a_y = after[2] if after else None
        w_y = ws[2] if ws else None
        adv = a_y - b_y if (a_y and b_y) else None
        ws_lines = round((adv - 18) / 18, 1) if adv else None
        print(f'{name:<14} {b_y!s:<12} {w_y!s:<12} {a_y!s:<12} {adv!s:<12} {ws_lines!s:<10}')


if __name__ == '__main__':
    main()
