# -*- coding: utf-8 -*-
"""
Measure how Word interprets `<w:line w:val="-N" w:lineRule="auto"/>` (negative
line spacing value with "auto" rule) — the OOXML pattern found in d4d126
paragraphs wi=289 (line=-200), wi=291 (line=-220), and likely wi=292.

Day 37 finding: Oxi's `--dump-layout` puts wi=289 (3 text fragments) all at
y=665.25 with cursor advance to next paragraph only 7.7pt, while Word renders
the same paragraph occupying 27.5pt (clearly 2 lines). The OOXML has
`<w:line w:val="-200" w:lineRule="auto"/>` + `<w:sz w:val="18"/>` (= 9pt)
+ a bracketPair drawing in R0.

This script:
1. Opens d4d126
2. Finds paragraphs wi=289 ("当該職員の氏名"), wi=291 ("□　提供依頼申出者又は代理人"),
   wi=292 ("訪問可能時期")
3. For each, reports:
   - Format.LineSpacing (the resolved value, in pt)
   - Format.LineSpacingRule (wdLineSpaceSingle=0, wdLineSpaceAtLeast=3,
     wdLineSpaceExactly=4, wdLineSpaceMultiple=5)
   - Per-character Y position to count line wraps and line height
4. Compares to a reference paragraph (e.g., wi=288 with positive line=auto)
   for contrast.

Output: pipeline_data/d4d126_neg_line_value.json

Interpretation:
- If Word's resolved LineSpacing == |line| / 20 pt → Word reads negative as
  "exact" magnitude regardless of `lineRule="auto"`
- If Word's resolved LineSpacing == some auto-computed value → Word treats
  it as auto anyway (ignores negative sign)
- If LineSpacingRule == 4 (exact) → Word interprets negative w:line as
  "exact" mode override
"""
import os
import sys
import json

sys.stdout.reconfigure(encoding='utf-8')

import win32com.client as wc

DOC = os.path.abspath('tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx')
OUT = os.path.abspath('pipeline_data/d4d126_neg_line_value.json')

WD_HPOS = 5
WD_VPOS = 6

# wdLineSpacing constants
WD_LS_NAMES = {
    0: 'wdLineSpaceSingle',
    1: 'wdLineSpace1pt5',
    2: 'wdLineSpaceDouble',
    3: 'wdLineSpaceAtLeast',
    4: 'wdLineSpaceExactly',
    5: 'wdLineSpaceMultiple',
}

TARGETS = [
    ('wi_289', '当該職員の氏名'),
    ('wi_291', '提供依頼申出者又は代理人'),
    ('wi_292', '訪問可能時期'),
    ('wi_288_reference', '日本政府の職員が提供依頼'),
]


def find_para_by_prefix(d, prefix):
    n = d.Paragraphs.Count
    for i in range(1, n + 1):
        txt = d.Paragraphs(i).Range.Text or ''
        if prefix in txt:
            return i
    return None


def measure_para(d, target_pi, label):
    para = d.Paragraphs(target_pi)
    rng = para.Range
    full_text = rng.Text.rstrip('\r\n\x07')
    pf = para.Format
    fnt = rng.Font

    ls_rule_int = pf.LineSpacingRule
    ls_rule_name = WD_LS_NAMES.get(ls_rule_int, f'unknown({ls_rule_int})')

    result = {
        'label': label,
        'para_index_1based': target_pi,
        'text': full_text[:120],
        'text_len': len(full_text),
        'line_spacing_pt': pf.LineSpacing,
        'line_spacing_rule_int': ls_rule_int,
        'line_spacing_rule_name': ls_rule_name,
        'space_before_pt': pf.SpaceBefore,
        'space_after_pt': pf.SpaceAfter,
        'font_size_pt': fnt.Size,
        'font_name': fnt.NameFarEast or fnt.Name,
    }

    # Per-char Y to count lines
    start = rng.Start
    ys = []
    for i in range(min(len(full_text), 100)):
        ch_rng = d.Range(start + i, start + i)
        y = ch_rng.Information(WD_VPOS)
        ys.append(round(y, 2))

    unique_ys = sorted(set(ys))
    result['line_count'] = len(unique_ys)
    result['line_ys'] = unique_ys
    if len(unique_ys) >= 2:
        deltas = [unique_ys[i + 1] - unique_ys[i] for i in range(len(unique_ys) - 1)]
        result['line_y_deltas'] = [round(d, 2) for d in deltas]
        result['avg_line_y_delta_pt'] = round(sum(deltas) / len(deltas), 2)

    print(f'\n=== {label} (para #{target_pi}) ===')
    print(f'  text[:60]={full_text[:60]!r}')
    print(f'  fs={result["font_size_pt"]} font={result["font_name"]}')
    print(f'  LineSpacing(pt)={result["line_spacing_pt"]:.3f}')
    print(f'  LineSpacingRule={ls_rule_name} ({ls_rule_int})')
    print(f'  SpaceBefore={result["space_before_pt"]} SpaceAfter={result["space_after_pt"]}')
    print(f'  line_count={result["line_count"]}')
    print(f'  line_ys={result["line_ys"][:8]}')
    if 'avg_line_y_delta_pt' in result:
        print(f'  line_y_deltas={result["line_y_deltas"][:6]}')
        print(f'  avg_line_y_delta_pt={result["avg_line_y_delta_pt"]}')

    return result


def measure():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(DOC, ReadOnly=True)
    out = {'doc': os.path.basename(DOC), 'paragraphs': []}
    try:
        for label, prefix in TARGETS:
            pi = find_para_by_prefix(d, prefix)
            if pi is None:
                print(f'!! {label}: prefix {prefix!r} not found')
                continue
            out['paragraphs'].append(measure_para(d, pi, label))
    finally:
        d.Close(SaveChanges=False)
        word.Quit()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f'\nWrote {OUT}')


if __name__ == '__main__':
    measure()
