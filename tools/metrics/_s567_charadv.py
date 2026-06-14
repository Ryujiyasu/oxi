# -*- coding: utf-8 -*-
"""Measure Word's per-character horizontal advance for harassmanual P18 (HGS
font), via Information(wdHorizontalPositionRelativeToPage) per char. Compare to
Oxi's proportional kana widths."""
import os
import win32com.client as win32
WD_HPOS = 5  # wdHorizontalPositionRelativeToTextBoundary? use 5=wdHorizontalPositionRelativeToPage
DOCX = os.path.abspath("tools/golden-test/documents/docx/harassmanual_001466344.docx")
word = win32.gencache.EnsureDispatch("Word.Application"); word.Visible=False
doc = word.Documents.Open(DOCX, ReadOnly=True)
try:
    tgt=None
    for i in range(1,doc.Paragraphs.Count+1):
        if doc.Paragraphs(i).Range.Text.startswith("解決に時間を要する"): tgt=i;break
    rng=doc.Paragraphs(tgt).Range; t=rng.Text.rstrip("\r\n")
    # font name of first body run
    print("para font.Name=", rng.Font.Name, " NameFarEast=", rng.Font.NameFarEast)
    # x position of each char (only line 1, chars 0..33)
    xs=[]
    for k in range(0, 20):
        c=doc.Range(rng.Start+k, rng.Start+k+1)
        x=c.Information(5)  # wdHorizontalPositionRelativeToPage
        xs.append((t[k], x))
    print("Word per-char x (line1):")
    for i in range(len(xs)-1):
        ch,x=xs[i]; adv=xs[i+1][1]-x
        print(f"  {ch!r} x={x:.2f} adv={adv:.2f}")
finally:
    doc.Close(False); word.Quit()
