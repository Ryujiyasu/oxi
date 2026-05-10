"""Day 33 part 28 (2026-05-11) — Measure db9ca wi=37 per-character Y positions.

Hypothesis: Word allows first-line overflow for LRPB-tagged multi-line
paragraphs that cross page boundary. Test by extracting per-character
Y positions and finding where Word's actual line break is.
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOCX = 'tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx'


def main():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    abs_path = os.path.abspath(DOCX)
    d = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        # Get wi=37 paragraph
        p = d.Paragraphs(37)
        r = p.Range
        text = (r.Text or '')[:200]
        print(f'wi=37 text: {text!r}')
        print(f'wi=37 char count: {r.Characters.Count}')
        print()

        # Iterate each character, get its y position via collapsed Range
        char_count = r.Characters.Count
        prev_y = None
        prev_pg = None
        char_idx = 0
        line_breaks = []
        positions = []
        for i in range(1, char_count + 1):
            c = r.Characters(i)
            ch_range = d.Range(c.Start, c.Start)
            try:
                pg = int(ch_range.Information(3))
                y = round(ch_range.Information(6), 2)
                ch_text = (c.Text or '')
            except Exception as e:
                continue
            positions.append((i, pg, y, ch_text))
            if prev_y is None or pg != prev_pg or abs(y - prev_y) > 0.5:
                # New line or page
                ev = 'page' if (prev_pg is not None and pg != prev_pg) else 'line'
                print(f'  char[{i:>3}] pg={pg} y={y:>7} {ev}-start char={ch_text!r}')
                line_breaks.append((i, pg, y, ch_text))
            prev_y = y
            prev_pg = pg
            char_idx += 1
            if i > 350: break  # limit
        print()
        print(f'Total chars processed: {char_idx}, line/page breaks: {len(line_breaks)}')
    finally:
        d.Close(False)
        word.Quit()


if __name__ == '__main__':
    main()
