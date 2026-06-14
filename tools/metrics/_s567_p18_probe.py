# -*- coding: utf-8 -*-
"""Probe Word's wrap of the '解決に時間を要する' numbered-list para: font size,
display line count, and the character index where line 1 breaks (via
Information(10) wdFirstCharacterLineNumber scan). Disambiguates list-indent vs
charGrid as the wrap-width cause."""
import os
import win32com.client as win32
DOCX = os.path.abspath("tools/golden-test/documents/docx/harassmanual_001466344.docx")
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
try:
    target = None
    for i in range(1, doc.Paragraphs.Count+1):
        t = doc.Paragraphs(i).Range.Text
        if t.startswith("解決に時間を要する"):
            target = i; break
    rng = doc.Paragraphs(target).Range
    sz = rng.Font.Size
    txt = rng.Text.rstrip("\r\n")
    print(f"para#{target} font.size={sz}pt len={len(txt)} text={txt!r}")
    # left indent / first-line
    pf = rng.ParagraphFormat
    print(f"LeftIndent={pf.LeftIndent:.1f}pt FirstLineIndent={pf.FirstLineIndent:.1f}pt")
    # line-number scan: line number of each char
    base = None; breaks = []
    prev_ln = None
    for k in range(len(txt)):
        c = doc.Range(rng.Start+k, rng.Start+k+1)
        ln = c.Information(10)  # wdFirstCharacterLineNumber
        if base is None: base = ln
        if prev_ln is not None and ln != prev_ln:
            breaks.append((k, txt[max(0,k-1)], txt[k]))
        prev_ln = ln
    nlines = (prev_ln - base + 1) if prev_ln is not None else 1
    print(f"display_lines={nlines}")
    for (k, before, after) in breaks:
        print(f"  line-break at char {k}: ...{txt[max(0,k-12):k]!r} | {txt[k:k+12]!r}")
finally:
    doc.Close(False); word.Quit()
