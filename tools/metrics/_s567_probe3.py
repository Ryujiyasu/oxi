# -*- coding: utf-8 -*-
"""Measure a SPECIFIC paragraph (by text prefix) accurately: PageSetup colW,
actual LeftIndent/FirstLineIndent, columns, line-break char counts."""
import os, sys
import win32com.client as win32
DOCX=os.path.abspath(sys.argv[1]); PREFIX=sys.argv[2]
word=win32.gencache.EnsureDispatch("Word.Application"); word.Visible=False
doc=word.Documents.Open(DOCX,ReadOnly=True)
try:
    tgt=None
    for i in range(1,doc.Paragraphs.Count+1):
        if doc.Paragraphs(i).Range.Text.startswith(PREFIX): tgt=i;break
    rng=doc.Paragraphs(tgt).Range; t=rng.Text.rstrip("\r\n"); pf=rng.ParagraphFormat
    ps=doc.PageSetup
    cols=doc.Sections(rng.Sections(1).Index).PageSetup.TextColumns
    ncol=cols.Count; colw=cols(1).Width if ncol>=1 else None
    print(f"para#{tgt} len={len(t)} sz={rng.Font.Size} Left={pf.LeftIndent:.1f} First={pf.FirstLineIndent:.1f}")
    print(f"PageWidth={ps.PageWidth:.1f} L={ps.LeftMargin:.1f} R={ps.RightMargin:.1f} colW(full)={ps.PageWidth-ps.LeftMargin-ps.RightMargin:.1f}")
    print(f"TextColumns count={ncol} col1.Width={colw}")
    base=None;prev=None;breaks=[]
    for k in range(len(t)):
        ln=doc.Range(rng.Start+k,rng.Start+k+1).Information(10)
        if base is None:base=ln
        if prev is not None and ln!=prev:breaks.append(k)
        prev=ln
    print(f"lines={(prev-base+1) if prev else 1} breaks={breaks}")
    eff = (colw if colw else ps.PageWidth-ps.LeftMargin-ps.RightMargin)
    if len(breaks)>=2:
        l2=breaks[1]-breaks[0]; cw=eff-pf.LeftIndent
        print(f"line2={l2}ch cont_w={cw:.1f} eff_charW={cw/l2:.2f}pt")
finally:
    doc.Close(False);word.Quit()
