# -*- coding: utf-8 -*-
"""S492h — test the OIKOMI hypothesis: Word compresses L1 (demand) to pull in a
line-start-PROHIBITED char (、。）」 etc.) that would otherwise be stranded at L2 head;
breaks at natural when the wrap char is freely wrappable. For b837's over-full paras
(idx 29) vs fit-at-natural paras (idx 66/71/73), report each Word line's last char and
the NEXT line's first char + whether that first char is line-start-prohibited. If the
over-full (compressing) paras have prohibited chars at the would-be wrap and the
fit-natural ones don't, oikomi is the gate. cp932-safe (UTF-8 file, ASCII verdict).
"""
import os, glob, json, re
import win32com.client as w32

WD_VPOS = 6
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/b837*.docx')[0])
rows = json.load(open('c:/tmp/b837_align.json', encoding='cp932'))
targets = {29: 'OVER-FULL(compress)', 66: 'fit-natural', 71: 'fit-natural', 73: 'fit-natural',
           23: 'jc=left', 69: 'jc=left'}
pref = {i: rows[i]['word']['norm'][:14] for i in targets if i < len(rows) and rows[i]['oxi']}

PROH = set('）〕］｝〉》」』】〙〗)]}、。，．：；？！・‥…ー')  # line-start-prohibited (Word default)

word = w32.DispatchEx('Word.Application'); word.Visible = False
try:
    wdoc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        p2i = {v: k for k, v in pref.items()}
        for p in wdoc.Paragraphs:
            clean = p.Range.Text.replace('\r', '').replace('\x07', '').replace('\n', '')
            k = re.sub(r'\s', '', clean)[:14]
            if k not in p2i:
                continue
            idx = p2i[k]
            rng = p.Range; txt = rng.Text; start = rng.Start
            # build (char -> line index) by y
            y0 = wdoc.Range(start, start).Information(WD_VPOS)
            seq = []
            cur_line = 0; prev_y = y0
            for i in range(len(txt)):
                ch = txt[i]
                if ch in ('\r', '\n', '\x07'):
                    continue
                y = wdoc.Range(start + i, start + i).Information(WD_VPOS)
                if y > prev_y + 2:
                    cur_line += 1; prev_y = y
                seq.append((ch, cur_line))
            nlines = seq[-1][1] + 1 if seq else 0
            print("\nidx %d [%s] %d lines:" % (idx, targets[idx], nlines))
            for ln in range(nlines):
                chs = [c for c, l in seq if l == ln]
                last = chs[-1] if chs else ''
                nxt = next((c for c, l in seq if l == ln + 1), '')
                proh = nxt in PROH
                print("   L%d last=%r  next-line-first=%r%s" %
                      (ln, last, nxt, '  <PROHIBITED at L-head -> oikomi candidate' if proh else ''))
    finally:
        wdoc.Close(False)
finally:
    word.Quit()
