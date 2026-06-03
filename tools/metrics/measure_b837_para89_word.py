"""S492c — GROUND TRUTH for b837 para89 (resolve the contradictions). Robust substring
match (Word Find) to locate the paragraph, then report: indent (first-line X, cont X),
line count, per-line char counts, and per-char X on the longest line (to detect sub-12pt
compression). Decides: is Oxi's 538.9>524.4 overflow a real over-pack, and does Word fit
3 lines via sub-grid compression (would contradict R35 'CJK never <12.0')?
"""
import os, glob, re
import win32com.client as w32

WD_VPOS, WD_HPOS = 6, 5
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/b837*.docx')[0])
NEEDLE = '公開が望まれる'

word = w32.DispatchEx('Word.Application'); word.Visible = False
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        target = None
        for p in doc.Paragraphs:
            if NEEDLE in p.Range.Text:
                target = p; break
        if target is None:
            print("NEEDLE not found in any paragraph; searching full doc text...")
            full = doc.Content.Text
            idx = full.find(NEEDLE)
            print("found in doc.Content at char", idx if idx >= 0 else "NOT FOUND")
        else:
            rng = target.Range; txt = rng.Text; start = rng.Start
            # alignment + indent
            try:
                al = target.Alignment
                li = target.LeftIndent; fli = target.FirstLineIndent
            except Exception:
                al = li = fli = '?'
            print("para alignment=%s leftIndent=%s firstLineIndent=%s" % (al, li, fli))
            print("text[:30]=%r  len(clean)=%d" % (txt[:30], len(txt.replace('\r','').replace('\x07',''))))
            y0 = doc.Range(start, start).Information(WD_VPOS)
            rows = []  # per char (ch, x, y)
            for i in range(len(txt)):
                ch = txt[i]
                if ch in ('\r', '\n', '\x07'):
                    continue
                x = doc.Range(start + i, start + i).Information(WD_HPOS)
                y = doc.Range(start + i, start + i).Information(WD_VPOS)
                rows.append((ch, round(x, 2), round(y, 2)))
            # group into lines by y
            from collections import OrderedDict
            lines = OrderedDict()
            for ch, x, y in rows:
                lines.setdefault(y, []).append((ch, x))
            print("Word para89: %d lines" % len(lines))
            for li_i, (y, chs) in enumerate(lines.items()):
                xs = [c[1] for c in chs]
                x0 = min(xs); xn = max(xs)
                # advance estimate: (xn - x0)/(n-1)
                n = len(chs)
                adv = (xn - x0) / (n - 1) if n > 1 else 0
                print("  line %d y=%.1f: %d chars, x0=%.1f xlast=%.1f avg_adv=%.2f%s"
                      % (li_i + 1, y, n, x0, xn, adv, '  <<SUB-12pt' if 0 < adv < 11.7 else ''))
    finally:
        doc.Close(False)
finally:
    word.Quit()
