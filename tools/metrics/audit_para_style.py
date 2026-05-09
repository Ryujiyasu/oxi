"""Day 32 part 9 — Paragraph style audit.

Day 32 part 8 left H3 (paragraph style) as the next pending hypothesis
for Bug 2 conditional detector. This tool extracts paragraph styles
for the first 5 paragraphs of each Class A doc + preserve sample.

Hypothesis: Word skips centering for Heading-style paragraphs even
when docGrid is set. Class A first-paragraph titles may be Heading
style, while preserve-class middle paragraphs are Normal style.

Output: style name + isHeading flag + style-level snap setting.
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX_DIR = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx')


def find_docx(doc_id):
    for f in os.listdir(DOCX_DIR):
        if f.startswith(doc_id) and f.endswith('.docx'):
            return os.path.join(DOCX_DIR, f)
    return None


def audit(doc_id, label, n_paras=5):
    import win32com.client as wc
    docx = find_docx(doc_id)
    if not docx:
        print(f'{label} {doc_id}: NOT FOUND')
        return
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
    try:
        n = min(d.Paragraphs.Count, n_paras)
        print(f'{label:<10} {doc_id}:')
        print(f'  {"i":>3} {"style":<25} {"snap":>4} {"y":>7} {"fs":>5} {"alignment":>9} {"linkedStyle":<18} text')
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            try:
                style_name = str(p.Style.NameLocal)
            except Exception:
                style_name = '?'
            try:
                linked = ''
                if hasattr(p.Style, 'LinkStyle') and p.Style.LinkStyle:
                    linked = str(p.Style.LinkStyle.NameLocal)
            except Exception:
                linked = ''
            try:
                snap = p.Format.SnapToGrid
            except Exception:
                snap = -1
            try:
                fs = r.Font.Size
            except Exception:
                fs = -1
            try:
                ta = p.Format.TextAlignment  # wdAlignVerticalBaseline=4, etc.
            except Exception:
                ta = -1
            text = (r.Text or '').strip()[:40]
            print(f'  {i:>3} {style_name[:25]!s:<25} {snap:>4} {round(cr.Information(6), 2):>7.2f} {fs:>5} {ta:>9} {linked[:18]!s:<18} {text!r}')
    finally:
        d.Close(False)
        word.Quit()


def main():
    print('=== Class A docs ===')
    class_a = ['bd90b00ab7a7', 'de6e32b5960b', 'db9ca18368cd', 'd77a58485f16']
    for d in class_a:
        audit(d, 'Class A')
        print()

    print('=== Preserve sample ===')
    preserve = ['e3c545fac7a7', '0e7af1ae8f21']
    for d in preserve:
        audit(d, 'Preserve')
        print()


if __name__ == '__main__':
    main()
