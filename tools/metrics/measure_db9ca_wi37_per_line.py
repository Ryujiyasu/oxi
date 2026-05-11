"""Day 33 part 30 (2026-05-11) — Per-line measurement of db9ca wi=37 via
Range.Information(wdFirstCharacterLineNumber).

Goal: identify EXACT y position where Word draws each line of wi=37, to
determine if Word allows first-line descender intrusion into bottom margin.
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOCX = 'tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx'

# Word constants
wdFirstCharacterLineNumber = 10
wdHorizontalPositionRelativeToPage = 5
wdVerticalPositionRelativeToPage = 6
wdActiveEndPageNumber = 3


def main():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    abs_path = os.path.abspath(DOCX)
    d = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        p = d.Paragraphs(37)
        r = p.Range
        char_count = r.Characters.Count
        print(f'wi=37 char count: {char_count}')

        # For each character, get its line number and y position
        cur_line_y = None
        cur_line_pg = None
        cur_line_num = None
        cur_line_start_char = 1
        lines = []
        for i in range(1, char_count + 1):
            c = r.Characters(i)
            ch_range = d.Range(c.Start, c.Start)
            try:
                line_num = int(ch_range.Information(wdFirstCharacterLineNumber))
                pg = int(ch_range.Information(wdActiveEndPageNumber))
                y = round(ch_range.Information(wdVerticalPositionRelativeToPage), 2)
                ch_text = c.Text or ''
            except Exception as e:
                continue
            if cur_line_num is None or line_num != cur_line_num or pg != cur_line_pg:
                if cur_line_num is not None:
                    lines.append({
                        'line_num': cur_line_num, 'pg': cur_line_pg, 'y': cur_line_y,
                        'start_char': cur_line_start_char, 'end_char': i - 1,
                    })
                cur_line_num = line_num
                cur_line_pg = pg
                cur_line_y = y
                cur_line_start_char = i
        # Flush last line
        if cur_line_num is not None:
            lines.append({
                'line_num': cur_line_num, 'pg': cur_line_pg, 'y': cur_line_y,
                'start_char': cur_line_start_char, 'end_char': char_count,
            })

        print(f'\nwi=37 lines: {len(lines)}')
        for L in lines:
            text_chars = r.Characters
            first_chars = ''.join(text_chars(i).Text for i in range(L['start_char'], min(L['start_char'] + 30, L['end_char'] + 1)))
            print(f'  line {L["line_num"]:>3} pg={L["pg"]} y={L["y"]:>6} chars[{L["start_char"]}..{L["end_char"]}] | {first_chars!r}')
    finally:
        d.Close(False)
        word.Quit()


if __name__ == '__main__':
    main()
