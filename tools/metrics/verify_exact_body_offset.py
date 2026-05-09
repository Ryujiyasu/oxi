"""Day 32 part 13 — Verify Word body lh=exact text position via per-char COM.

Hypothesis (Day 32 part 12): Word body paragraphs with lh=exact place
text at line BOX TOP (offset=0), while Oxi mod.rs:5393 places at line
BOX BOTTOM (offset=lh-fs).

COM Information(6) on cr.Start returns paragraph anchor (line box top).
But to verify TEXT GLYPH position within the line box, we need to
measure a specific character's render y.

Word's Information(WD_VPOS=6) on Range(start, start+1) returns the y
of that character. If Word places text at line top:
  char_y = paragraph_anchor + 0  (Information(6) on cr_start = char_y)
If Word places text at line bottom:
  char_y = paragraph_anchor + (lh - fs)

Test: measure Information(6) on cr_start vs Information(6) on cr_first_char.
If both are equal → text at top.
If first_char is offset by (lh-fs) → text at bottom.

Tested on bd90b00 pi=11 (existing doc, lh=exact 16pt fs=11.5pt).
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
        # Test paragraphs:
        # pi=11: body, lh=Exact 16pt, fs=11.5pt → expected offset = 4.5pt if at bottom
        # pi=42: body, lh=Exact 10pt, fs=8pt → expected offset = 2pt if at bottom
        # pi=2: body, lh=Single 12pt, fs=14pt → reference (no exact mode)
        # pi=13: body, lh=Exact 11pt, fs=10.5pt → expected offset = 0.5pt if at bottom
        targets = [2, 11, 13, 42]
        print(f'{"i":>3} {"start_y":>8} {"char1_y":>8} {"diff":>6} {"lh_rule":>7} {"lh_val":>6} {"fs":>5} text')
        for i in targets:
            p = d.Paragraphs(i)
            r = p.Range
            cr_start = d.Range(r.Start, r.Start)
            # Find first non-whitespace character offset
            text = r.Text or ''
            char_offset = 0
            for ch in text:
                if ch.strip():
                    break
                char_offset += 1
            # Skip leading whitespace, get first content char
            cr_char = d.Range(r.Start + char_offset, r.Start + char_offset + 1)
            try:
                start_y = round(cr_start.Information(6), 2)
            except Exception:
                start_y = -1
            try:
                char_y = round(cr_char.Information(6), 2)
            except Exception:
                char_y = -1
            try:
                lh_rule = p.Format.LineSpacingRule
                lh_val = p.Format.LineSpacing
            except Exception:
                lh_rule, lh_val = -1, -1
            try:
                fs = r.Font.Size
            except Exception:
                fs = -1
            text_short = text.strip()[:40]
            print(f'{i:>3} {start_y:>8.2f} {char_y:>8.2f} {char_y-start_y:>+6.2f} {lh_rule:>7} {lh_val:>6.1f} {fs:>5} {text_short!r}')

        print('\nInterpretation:')
        print('  diff = 0  → text at line box top (Word offset = 0)')
        print('  diff > 0  → text below line box top by that amount')
        print('  for lh=exact + fs case, diff matches (lh-fs) means text at bottom')

    finally:
        d.Close(False)
        word.Quit()


if __name__ == '__main__':
    main()
