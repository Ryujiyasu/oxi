"""Day 33 part 18 — Measure last-page-1 paragraph y for FP_A/B/C variants.

Hypothesis test: if framePr-wrapped footer paragraphs are excluded from
inline footer_h, then FP_A (empty only), FP_B (framePr+empty), FP_C (framePr only)
should all have same body content area on page 1, hence same last-fitting paragraph.
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOCS = [
    'tools/golden-test/repros/frame_pr_footer/FP_A_empty_only.docx',
    'tools/golden-test/repros/frame_pr_footer/FP_B_framepr_plus_empty.docx',
    'tools/golden-test/repros/frame_pr_footer/FP_C_framepr_only.docx',
]


def measure(docx_path):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    abs_path = os.path.abspath(docx_path)
    d = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        n = d.Paragraphs.Count
        last_p1 = None
        first_p2 = None
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            try:
                pg = int(cr.Information(3))
                y = cr.Information(6)
            except: continue
            text = (r.Text or '').strip()
            if pg == 1:
                last_p1 = (i, y, text[:20])
            elif pg == 2 and first_p2 is None:
                first_p2 = (i, y, text[:20])
                break
        return last_p1, first_p2, n
    finally:
        d.Close(False)
        word.Quit()


def main():
    print(f'{"variant":<35} {"last_p1":<35} {"first_p2":<35}')
    for d in DOCS:
        name = os.path.basename(d)
        try:
            l, f, _ = measure(d)
        except Exception as e:
            print(f'{name}: error {e}'); continue
        l_s = f'i={l[0]} y={l[1]:.2f} {l[2]}' if l else '-'
        f_s = f'i={f[0]} y={f[1]:.2f} {f[2]}' if f else '-'
        print(f'{name:<35} {l_s:<35} {f_s:<35}')


if __name__ == '__main__':
    main()
