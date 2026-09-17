# -*- coding: utf-8 -*-
"""Where do Word and Oxi break the SAME paragraph differently, and what is on that line?

The census (`_pb_oikomi_census.py`) records Word's own line breaks per paragraph; the GDI
renderer's `--dump-layout` records Oxi's. Joining them by line text turns "Oxi packs one
character too many somewhere in this document" into a labelled set: every line where the two
engines disagree, with the punctuation that line carries and the character each engine chose
to break before. Four hypotheses have already died on single paragraphs (empty-paragraph
height, the table-exit addend, "compress only when it saves a line", per-mark caps), so this
stops proposing rules and starts from the disagreements.

Join: Word's lines for a paragraph are consecutive and cover the paragraph exactly, and so
are Oxi's, so the two sequences are aligned by walking Oxi's rendered lines until the
concatenation matches the paragraph's Word text. A paragraph is reported when any line's
character count differs.

Readout per disagreeing paragraph: each side's line lengths, the first line where they part,
and for that line the trailing context of both, plus the marks each side counted.
"""
import json, sys, os
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
MARKS = '、。，．・：；（）「」［］〕〔【】'
CENSUS = sys.argv[1] if len(sys.argv) > 1 else 'pipeline_data/oikomi_census/golden4_v3.jsonl'
DUMPDIR = 'pipeline_data/oikomi_census'


def oxi_lines(path):
    """Rendered lines in reading order: (y-ordered within a page, pages in order)."""
    d = json.load(open(path, encoding='utf-8'))
    out = []
    for page in d['pages']:
        rows = defaultdict(list)
        for e in page['elements']:
            if e['type'] == 'text' and e.get('text'):
                rows[round(e['y'], 2)].append(e)
        for y in sorted(rows):
            es = sorted(rows[y], key=lambda e: e['x'])
            out.append(''.join(e['text'] for e in es))
    return out


def main():
    byparagraph = defaultdict(list)
    for line in open(CENSUS, encoding='utf-8'):
        line = line.strip()
        if not line.startswith('{'):
            continue
        r = json.loads(line)
        if 'error' in r or 'text' not in r:
            continue
        byparagraph[(r['doc'], r['para'])].append(r)

    dumps = {}
    n_para = n_diff = 0
    for (doc, para), rows in sorted(byparagraph.items()):
        rows.sort(key=lambda r: r['line'])
        stem = os.path.splitext(doc)[0]
        if stem not in dumps:
            p = os.path.join(DUMPDIR, f'oxi_{stem}.json')
            dumps[stem] = oxi_lines(p) if os.path.exists(p) else []
        ox = dumps[stem]
        want = ''.join(r['text'] for r in rows)
        key = rows[0]['text'][:10]
        start = next((i for i, l in enumerate(ox) if l.startswith(key)), None)
        if start is None:
            continue
        got, j = '', start
        mine = []
        while j < len(ox) and len(got) < len(want):
            mine.append(ox[j]); got += ox[j]; j += 1
        if got.replace(' ', '') != want.replace(' ', ''):
            continue
        n_para += 1
        wlen = [len(r['text']) for r in rows]
        olen = [len(l) for l in mine]
        if wlen == olen:
            continue
        n_diff += 1
        k = next((i for i in range(min(len(wlen), len(olen))) if wlen[i] != olen[i]), 0)
        wl = rows[k]['text'] if k < len(rows) else ''
        ol = mine[k] if k < len(mine) else ''
        print(f'{doc} para={para} word={wlen} oxi={olen} first_diff_line={k}')
        print(f'   word[{len(wl):3}] …{wl[-14:]!r}  marks={sum(1 for c in wl if c in MARKS)}')
        print(f'   oxi [{len(ol):3}] …{ol[-14:]!r}  marks={sum(1 for c in ol if c in MARKS)}')
        extra = ol[len(wl):] if len(ol) > len(wl) else ''
        print(f'   oxi kept extra={extra!r}  word broke before={wl and want[len(wl)] if len(want) > len(wl) else ""!r}')
    print(f'--- paragraphs joined={n_para} disagreeing={n_diff}')


main()
