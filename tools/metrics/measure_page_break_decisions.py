"""Day 33 part 54 — Page-break decision side-by-side comparison (R7.6).

For each 备考 cluster doc:
- Word: walk paragraphs in order, find page transitions (i where pg changes)
- Identify "last paragraph on page N" and "first paragraph on page N+1"
- Word's decision: last fits at (lastPara.y, lastPara.y + line_h ≤ page_bot)
- Oxi: from OXI_DUMP_BREAK output, find the actual break decision (which paragraph)
- Compare: is Oxi's break paragraph EARLIER than Word's? (= Oxi too conservative)

Also: for each Oxi page-end paragraph (last fitting on page), compute margin =
page_bot - (cursor_y + line_h). If margin is tiny (0-2pt), the page-break is
marginal and small drift can flip it.

Output: pipeline_data/page_break_decisions_<doc>.csv
       per Word page transition: word_last_i, word_last_text, oxi_break_i, oxi_margin, drift.
"""
from __future__ import annotations
import os, sys, json, subprocess, re, csv
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc
import glob

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
RENDERER = os.path.abspath(os.path.join(REPO, 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe'))
PAGE_H = 841.95
WD_VPOS = 6
WD_PAGE = 3
WD_IN_TABLE = 12

CLUSTER = ['de6e32b5960b', 'd4d126dfe1d9', '6514f214e482',
           'a1d6e4efa2e7', '191cb5254cb2', '1636d28e2c46']


def measure_word(docx_path):
    """Walk paragraphs; return list with (i, page, y, in_table, text)."""
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    paras = []
    try:
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            try: y = round(cr.Information(WD_VPOS), 2)
            except: y = -1
            try: pg = int(cr.Information(WD_PAGE))
            except: pg = -1
            try: in_t = bool(r.Information(WD_IN_TABLE))
            except: in_t = False
            try: text = (r.Text or '').replace('\r', ' ').replace('\x07', '').strip()
            except: text = ''
            paras.append({'i': i, 'pg': pg, 'y': y, 'in_table': in_t, 'text': text[:35]})
    finally:
        d.Close(False)
        word.Quit()
    return paras


def get_oxi_break_dump(docx_path):
    """Run renderer with OXI_DUMP_BREAK, parse output."""
    log_path = r'C:\tmp\pb_dump.log'
    cmd = [RENDERER, os.path.abspath(docx_path), r'C:\tmp\pb_out',
           '--dump-layout=' + r'C:\tmp\pb_layout.json']
    env = dict(os.environ); env['OXI_DUMP_BREAK'] = '1'
    with open(log_path, 'w') as f:
        subprocess.run(cmd, stderr=f, stdout=subprocess.DEVNULL, env=env, timeout=180)
    breaks = []
    with open(log_path, encoding='utf-8') as f:
        for line in f:
            m = re.match(r'\[BR_DUMP\] pi=(\d+) line\d+ cursor_y=([\d.]+) eff_lh=([\d.]+) line_h=([\d.]+) sum=([\d.]+) pg_top=([\d.]+) pg_bot=([\d.]+) brk=(\w+) text=(.*)', line)
            if m:
                breaks.append({
                    'pi': int(m.group(1)),
                    'cursor_y': float(m.group(2)),
                    'eff_lh': float(m.group(3)),
                    'line_h': float(m.group(4)),
                    'sum': float(m.group(5)),
                    'pg_top': float(m.group(6)),
                    'pg_bot': float(m.group(7)),
                    'brk': m.group(8) == 'true',
                    'text': m.group(9).strip(),
                })
    return breaks


def find_page_transitions(paras):
    """Identify (last_on_page_N, first_on_page_N+1) pairs."""
    transitions = []
    prev_pg = None
    prev_p = None
    for p in paras:
        if p['pg'] < 1: continue
        if prev_pg is not None and p['pg'] != prev_pg:
            transitions.append({'last_on': prev_p, 'first_on': p})
        prev_pg = p['pg']
        prev_p = p
    return transitions


def main():
    for doc_id in CLUSTER:
        docx_path_list = glob.glob(f'tools/golden-test/documents/docx/{doc_id}*')
        if not docx_path_list:
            print(f'{doc_id}: NOT FOUND')
            continue
        docx_path = docx_path_list[0]
        print(f'\n=== {doc_id} ===')
        word_paras = measure_word(docx_path)
        word_transitions = find_page_transitions(word_paras)
        oxi_breaks = get_oxi_break_dump(docx_path)

        # For each Word page transition, find the break point and Oxi's margin
        print(f'  Word transitions: {len(word_transitions)}')
        print(f'  Oxi BR_DUMP records: {len(oxi_breaks)}')

        # Oxi breaks where brk=true mean Oxi pushed paragraph to next page
        oxi_brk_true = [b for b in oxi_breaks if b['brk']]
        oxi_brk_false_margin = [(b['pg_bot'] - b['sum']) for b in oxi_breaks if not b['brk']]
        if oxi_brk_false_margin:
            avg_margin = sum(oxi_brk_false_margin) / len(oxi_brk_false_margin)
            min_margin = min(oxi_brk_false_margin)
            print(f'  Oxi non-break paragraphs: avg margin to pg_bot = {avg_margin:.1f}pt, min = {min_margin:.1f}pt')
        print(f'  Oxi page break events: {len(oxi_brk_true)}')

        # Show each Oxi break event with margin
        for b in oxi_brk_true:
            overflow = b['sum'] - b['pg_bot']
            print(f'    pi={b["pi"]:>3} cursor_y={b["cursor_y"]:>7.2f} sum={b["sum"]:>7.2f} '
                  f'pg_bot={b["pg_bot"]:>7.2f} overflow={overflow:+.2f}pt text={b["text"][:30]!r}')


if __name__ == '__main__':
    main()
