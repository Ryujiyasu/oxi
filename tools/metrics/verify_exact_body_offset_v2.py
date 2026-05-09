"""Day 32 part 13b — Verify text position via multiple WdInformation values.

Day 32 part 13a used Information(6) but it returns paragraph anchor for
both para start and first char. This v2 tries all relevant WdInformation
values to find one that returns the text glyph position.

WdInformation values:
  1 = ActiveEndAdjustedPageNumber
  3 = ActiveEndPageNumber
  5 = HorizontalPositionRelativeToPage
  6 = VerticalPositionRelativeToPage
  7 = HorizontalPositionRelativeToTextBoundary
  8 = VerticalPositionRelativeToTextBoundary
  9 = FirstCharacterColumnNumber
 10 = FirstCharacterLineNumber
 12 = WithInTable

Test on bd90b00 pi=11 (lh=Exact 16, fs=11.5) — expected difference if
exact-mode places text at bottom: diff = 4.5pt.
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')


def main():
    import win32com.client as wc
    docx = r'C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\bd90b00ab7a7_order_05.docx'
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
    try:
        targets = [2, 11, 13, 42]
        info_codes = [3, 5, 6, 7, 8, 9, 10]

        for i in targets:
            p = d.Paragraphs(i)
            r = p.Range
            cr_start = d.Range(r.Start, r.Start)
            text = r.Text or ''
            char_offset = 0
            for ch in text:
                if ch.strip():
                    break
                char_offset += 1
            # First content char
            cr_char = d.Range(r.Start + char_offset, r.Start + char_offset + 1)
            try:
                lh_rule = p.Format.LineSpacingRule
                lh_val = p.Format.LineSpacing
            except Exception:
                lh_rule, lh_val = -1, -1
            try:
                fs = r.Font.Size
            except Exception:
                fs = -1
            print(f'\n=== pi={i} lh_rule={lh_rule} lh_val={lh_val} fs={fs} text={text.strip()[:40]!r} ===')
            for code in info_codes:
                try:
                    s = round(cr_start.Information(code), 3)
                except Exception:
                    s = '?'
                try:
                    c = round(cr_char.Information(code), 3)
                except Exception:
                    c = '?'
                names = {
                    3: 'PageNumber', 5: 'HorizPos', 6: 'VertPos',
                    7: 'HorizPosBoundary', 8: 'VertPosBoundary',
                    9: 'FirstCharCol', 10: 'FirstCharLine',
                }
                marker = ''
                try:
                    if isinstance(s, (int, float)) and isinstance(c, (int, float)):
                        diff = round(c - s, 3)
                        if abs(diff) > 0.001:
                            marker = f' <<DIFF: {diff:+.3f}>>'
                except Exception:
                    pass
                print(f'  Info({code} {names.get(code, "?"):<18}): start={s} char={c}{marker}')

        # Also try BoundingBox if available (likely not)
        print('\n--- Try Range methods ---')
        p = d.Paragraphs(11)
        r = p.Range
        print(f'  Range.Start = {r.Start}, Range.End = {r.End}')
        # Try selecting the range and getting Selection info
        try:
            r.Select()
            sel = word.Selection
            print(f'  Selection vertical position (after Select): {sel.Information(6)}')
        except Exception as e:
            print(f'  Selection method error: {e}')

    finally:
        d.Close(False)
        word.Quit()


if __name__ == '__main__':
    main()
