# -*- coding: utf-8 -*-
"""
Measure per-character widths for d4d126's w_i=245 paragraph.

Background: Day 36 memory identifies pStyle "ac" (line=280 exact, sz=20 = 10pt
MS Mincho) as the root-cause class for d4d126's 4 of 5 page-delta mismatches.
w_i=245 is "○　法人等であって、その役員のうちに上記のいずれかに該当する者がある者"
(35 chars). Word renders it as 1 line; Oxi renders it as 2 lines. If Oxi's
char widths are slightly larger than Word's, accumulating across 33 chars,
the text would just barely overflow.

This script:
1. Opens d4d126
2. Finds paragraph w_i=245 (by partial text match for robustness)
3. For each character, captures HORIZONTAL_POSITION_RELATIVE_TO_PAGE
4. Outputs char widths (delta between consecutive X) + total width + line count

Output: pipeline_data/d4d126_w245_widths.json
"""
import os
import sys
import json

sys.stdout.reconfigure(encoding='utf-8')

import win32com.client as wc

DOC = os.path.abspath('tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx')
OUT = os.path.abspath('pipeline_data/d4d126_w245_widths.json')

# wdHorizontalPositionRelativeToPage = 5
# wdVerticalPositionRelativeToPage = 6
WD_HPOS = 5
WD_VPOS = 6
WD_FIRSTCHAR = 10  # wdFirstCharacterLineNumber alternative not used


def measure():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(DOC, ReadOnly=True)
    result = {}
    try:
        # Find paragraph by partial text — search for "○　法人等であって"
        # since that's the unique starting prefix
        target_prefix = '法人等であって'
        target_pi = None
        n = d.Paragraphs.Count
        for i in range(1, n + 1):
            txt = d.Paragraphs(i).Range.Text or ''
            if target_prefix in txt:
                target_pi = i
                break
        if target_pi is None:
            print('Target paragraph not found', file=sys.stderr)
            return
        print(f'Found target paragraph index: {target_pi}')
        para = d.Paragraphs(target_pi)
        rng = para.Range
        full_text = rng.Text.rstrip('\r\n\x07')
        print(f'Paragraph text (len={len(full_text)}):')
        print(repr(full_text))

        # Page setup
        ps = d.PageSetup
        result['page'] = {
            'pgH_pt': ps.PageHeight,
            'pgW_pt': ps.PageWidth,
            'topMargin_pt': ps.TopMargin,
            'bottomMargin_pt': ps.BottomMargin,
            'leftMargin_pt': ps.LeftMargin,
            'rightMargin_pt': ps.RightMargin,
        }

        # Get paragraph format info
        pf = para.Format
        fnt = rng.Font
        result['paragraph'] = {
            'para_index_1based': target_pi,
            'text': full_text,
            'text_len': len(full_text),
            'line_spacing_setting': pf.LineSpacing,
            'line_spacing_rule': pf.LineSpacingRule,
            'left_indent_pt': pf.LeftIndent,
            'first_line_indent_pt': pf.FirstLineIndent,
            'space_before_pt': pf.SpaceBefore,
            'space_after_pt': pf.SpaceAfter,
            'font_name': fnt.NameFarEast or fnt.Name,
            'font_size_pt': fnt.Size,
            'spacing': fnt.Spacing,
        }

        # Per-char measurement
        chars = []
        start = rng.Start
        end_excluding_marker = start + len(full_text)
        for i in range(len(full_text)):
            ch = full_text[i]
            ch_rng = d.Range(start + i, start + i)  # collapsed at i (left edge of char i)
            x = ch_rng.Information(WD_HPOS)
            y = ch_rng.Information(WD_VPOS)
            chars.append({
                'i': i,
                'ch': ch,
                'x_pt': round(x, 3),
                'y_pt': round(y, 3),
            })

        # Also measure right edge of last char
        end_rng = d.Range(end_excluding_marker, end_excluding_marker)
        end_x = end_rng.Information(WD_HPOS)
        end_y = end_rng.Information(WD_VPOS)

        # Compute per-char widths from consecutive X (same y only)
        widths = []
        for i in range(len(full_text)):
            if i == len(full_text) - 1:
                next_x = end_x
                next_y = end_y
            else:
                next_x = chars[i + 1]['x_pt']
                next_y = chars[i + 1]['y_pt']
            same_line = abs(next_y - chars[i]['y_pt']) < 1.0
            widths.append({
                'i': i,
                'ch': chars[i]['ch'],
                'x_pt': chars[i]['x_pt'],
                'y_pt': chars[i]['y_pt'],
                'width_pt': round(next_x - chars[i]['x_pt'], 3) if same_line else None,
                'wrapped_after': not same_line,
            })

        # Count distinct y values = line count
        unique_ys = sorted(set(c['y_pt'] for c in chars))
        result['line_count'] = len(unique_ys)
        result['line_ys'] = unique_ys
        result['end_pos'] = {'x_pt': round(end_x, 3), 'y_pt': round(end_y, 3)}
        result['widths'] = widths
        # Summary stats per line
        per_line = {}
        for w in widths:
            per_line.setdefault(w['y_pt'], []).append(w)
        result['per_line_summary'] = []
        for y in unique_ys:
            ws = per_line[y]
            real_widths = [w['width_pt'] for w in ws if w['width_pt'] is not None]
            total_w = sum(real_widths) if real_widths else 0
            result['per_line_summary'].append({
                'y_pt': y,
                'n_chars': len(ws),
                'total_width_pt': round(total_w, 3),
                'first_char_x': ws[0]['x_pt'],
                'last_char_x': ws[-1]['x_pt'],
            })

        print()
        print(f'Line count: {result["line_count"]}')
        for line in result['per_line_summary']:
            print(f'  y={line["y_pt"]:.2f}: {line["n_chars"]} chars, '
                  f'x={line["first_char_x"]:.2f}..{line["last_char_x"]:.2f}, '
                  f'total_w={line["total_width_pt"]:.2f}pt')
    finally:
        d.Close(SaveChanges=False)
        word.Quit()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'\nWrote {OUT}')


if __name__ == '__main__':
    measure()
