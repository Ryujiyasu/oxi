#!/usr/bin/env python3
"""char-budget cell-wrap: compare per-page LINE COUNT Oxi-dump vs Word-PDF.

Oxi: --dump-layout JSON (per-char text elements; group by (page, round(y)) = a line).
Word: ExportAsFixedFormat PDF (fitz rawdict; group spans by baseline y per page).

A page where Oxi line-count != Word line-count => a wrap/line-count divergence
(the char-budget cell-wrap symptom). Then drill with --lines to print each line.

Usage:
  _cb_cell_linecount.py <dump.json> <word.pdf> [--lines] [--page N]
"""
import sys, json, io

def oxi_lines(dump_path):
    d = json.load(io.open(dump_path, encoding='utf-8'))
    pages = []
    for pg in d['pages']:
        # bucket text elements by rounded y
        buckets = {}
        for e in pg['elements']:
            if e.get('type') != 'text':
                continue
            key = round(e['y'] * 2) / 2  # 0.5pt buckets
            buckets.setdefault(key, []).append(e)
        lines = []
        for y in sorted(buckets):
            els = sorted(buckets[y], key=lambda e: e['x'])
            txt = ''.join(decode(e['text']) for e in els)
            x0 = min(e['x'] for e in els)
            x1 = max(e['x'] + e['w'] for e in els)
            paras = sorted(set(e['para_idx'] for e in els))
            lines.append((y, x0, x1, txt, paras))
        pages.append(lines)
    return pages

def decode(s):
    try:
        return s.encode('latin1').decode('cp932')
    except Exception:
        return s

def word_lines(pdf_path):
    import fitz
    doc = fitz.open(pdf_path)
    pages = []
    for page in doc:
        d = page.get_text('rawdict')
        # collect chars with bbox
        rows = {}
        for block in d['blocks']:
            for line in block.get('lines', []):
                for span in line.get('spans', []):
                    for ch in span.get('chars', []):
                        bb = ch['bbox']
                        ybase = round(ch['origin'][1] * 2) / 2
                        rows.setdefault(ybase, []).append((bb[0], ch['c'], bb[2]))
        lines = []
        for y in sorted(rows):
            chs = sorted(rows[y])
            txt = ''.join(c for _, c, _ in chs)
            if not txt.strip():
                continue
            x0 = chs[0][0]; x1 = chs[-1][2]
            lines.append((y, x0, x1, txt))
        pages.append(lines)
    return pages

def main():
    dump, pdf = sys.argv[1], sys.argv[2]
    show = '--lines' in sys.argv
    only = None
    if '--page' in sys.argv:
        only = int(sys.argv[sys.argv.index('--page') + 1])
    ox = oxi_lines(dump)
    wd = word_lines(pdf)
    n = max(len(ox), len(wd))
    print(f"{'page':>4} {'oxi':>5} {'word':>5} {'diff':>5}")
    tot_o = tot_w = 0
    for i in range(n):
        o = len(ox[i]) if i < len(ox) else 0
        w = len(wd[i]) if i < len(wd) else 0
        tot_o += o; tot_w += w
        mark = '' if o == w else '  <-- DIFF'
        print(f"{i+1:>4} {o:>5} {w:>5} {o-w:>5}{mark}")
    print(f"{'SUM':>4} {tot_o:>5} {tot_w:>5} {tot_o-tot_w:>5}")
    if show and only:
        i = only - 1
        print(f"\n=== PAGE {only} OXI ({len(ox[i])} lines) ===")
        for y, x0, x1, txt, paras in ox[i]:
            print(f"  y={y:6.1f} x[{x0:6.1f},{x1:6.1f}] p{paras} {txt}")
        print(f"\n=== PAGE {only} WORD ({len(wd[i])} lines) ===")
        for y, x0, x1, txt in wd[i]:
            print(f"  y={y:6.1f} x[{x0:6.1f},{x1:6.1f}] {txt}")

if __name__ == '__main__':
    main()
