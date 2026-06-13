# -*- coding: utf-8 -*-
"""S559 wheel-(b) follow-up — find Oxi's COMPENSATING ERRORS. Word reserves the
default cellMar for ALL squeezed single cells, but applying it in Oxi cascades
{1:1323} because Oxi-OFF's 54/55 is held up by compensating line-count errors
(p19 under-counted 2 vs Word 3, balanced by an over-count nearby). This compares
per-paragraph LINE COUNTS Word (COM ComputeStatistics) vs Oxi-OFF (GDI dump) to
localize every discrepancy, grouped by Word page. The over-counts that balance
p19's under-count are the lever: fix them and cellMar applies universally.

Word line count: para.Range.ComputeStatistics(1) (wdStatisticLines) — 1 call/para.
Oxi-OFF line count: GDI --dump-layout, distinct y per (para_idx,cpi,cri,cci) group,
rendered with OXI_S559_DISABLE=1 (the pre-S559 baseline).
"""
import json
import os
import subprocess
import sys
import tempfile
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')
REPO = r'c:\Users\ryuji\oxi-main'
GDI = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
DOCX_REAL = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', '3a4f9fbe1a83_001620506.docx')
DOCX_COM = r'c:\tmp\3a4f9f.docx'


def norm(s):
    return ''.join(s.split())[:18]


def oxi_off_lines():
    env = dict(os.environ)
    env['OXI_S559_DISABLE'] = '1'
    with tempfile.TemporaryDirectory() as tmp:
        dump = os.path.join(tmp, 'l.json')
        subprocess.run([GDI, DOCX_REAL, os.path.join(tmp, 'p'), '--dump-layout=' + dump],
                       capture_output=True, env=env, timeout=300)
        d = json.load(open(dump, encoding='utf-8'))
    # per (para_idx,cpi,cri,cci): distinct y + text + page(min)
    ys = defaultdict(set)
    txt = {}
    pg = {}
    order = []
    for page in d['pages']:
        for e in page['elements']:
            if e.get('type') != 'text':
                continue
            k = (e.get('para_idx'), e.get('cell_para_idx'), e.get('cell_row_idx'), e.get('cell_col_idx'))
            if k not in txt:
                txt[k] = ''
                order.append(k)
            ys[k].add(round(e['y'], 1))
            if len(txt[k]) < 24:
                txt[k] += e.get('text', '')
            pg.setdefault(k, page['page'])
    out = []  # (norm_text, nlines, page) in document order
    for k in order:
        t = norm(txt[k])
        if t:
            out.append((t, len(ys[k]), pg[k]))
    return out


def word_lines():
    import win32com.client as w32
    word = w32.DispatchEx('Word.Application')
    word.Visible = False
    out = []
    try:
        wdoc = word.Documents.Open(os.path.abspath(DOCX_COM), ReadOnly=True)
        try:
            n = wdoc.Paragraphs.Count
            for i in range(1, n + 1):
                p = wdoc.Paragraphs(i).Range
                t = norm(p.Text or '')
                if not t:
                    continue
                nl = int(p.ComputeStatistics(1))  # wdStatisticLines
                pg = wdoc.Range(p.Start, p.Start).Information(3)  # collapsed-start page
                out.append((t, nl, int(pg)))
        finally:
            wdoc.Close(False)
    finally:
        word.Quit()
    return out


def main():
    print('rendering Oxi-OFF...')
    oxi = oxi_off_lines()
    print('  oxi paras:', len(oxi))
    print('measuring Word (ComputeStatistics)...')
    word = word_lines()
    print('  word paras:', len(word))

    # match by normalized text prefix (first occurrence), document order
    oxi_by_t = {}
    for t, nl, pg in oxi:
        oxi_by_t.setdefault(t, []).append((nl, pg))
    used = defaultdict(int)
    diffs = []  # (wpage, wtext, word_nl, oxi_nl, delta)
    for t, wnl, wpg in word:
        cand = oxi_by_t.get(t)
        if not cand:
            continue
        idx = used[t]
        if idx >= len(cand):
            idx = len(cand) - 1
        onl, opg = cand[idx]
        used[t] += 1
        if wnl != onl:
            diffs.append((wpg, t, wnl, onl, onl - wnl))

    # summary by page; focus region p17-p21 (p19 cascade origin)
    print('\n=== line-count discrepancies (Oxi-OFF vs Word), |delta|>=1 ===')
    print('total discrepant paras:', len(diffs))
    over = sum(1 for d in diffs if d[4] > 0)
    under = sum(1 for d in diffs if d[4] < 0)
    print('Oxi OVER-counts (Oxi>Word): %d   Oxi UNDER-counts: %d' % (over, under))
    net = sum(d[4] for d in diffs)
    print('net Oxi-Word line delta over matched paras: %+d' % net)

    print('\n=== region pages 17-22 (the p19 cascade origin) ===')
    for wpg, t, wnl, onl, dl in sorted(diffs):
        if 17 <= wpg <= 22:
            tag = 'OVER ' if dl > 0 else 'under'
            print('  p%-3d %s  Word=%d Oxi=%d (%+d)  %s' % (wpg, tag, wnl, onl, dl, t))

    # per-page net (where do over/under cancel?)
    page_net = defaultdict(int)
    for wpg, t, wnl, onl, dl in diffs:
        page_net[wpg] += dl
    print('\n=== per-page net line delta (Oxi-Word), nonzero pages ===')
    for pg in sorted(page_net):
        if page_net[pg] != 0:
            print('  p%-3d net=%+d' % (pg, page_net[pg]))


if __name__ == '__main__':
    main()
