# -*- coding: utf-8 -*-
"""Confirm the S567 charGrid spec on a 2nd real doc: measure a long CJK
paragraph's line-1 wrap char count, compute effective char width, compare to
linePitch vs default_fs vs body_fs."""
import os, sys
import win32com.client as win32
DOCX = os.path.abspath(sys.argv[1])
LP_TW = float(sys.argv[2]) if len(sys.argv)>2 else None
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
try:
    # pick the longest single-run-ish CJK paragraph (>40 chars, no table)
    best=None
    for i in range(1, doc.Paragraphs.Count+1):
        rng=doc.Paragraphs(i).Range
        t=rng.Text.rstrip("\r\n")
        if len(t)>=45 and rng.Tables.Count==0 and rng.Information(12)==False: # not in table
            best=i; break
    if best is None:
        print("no long para found"); sys.exit(0)
    rng=doc.Paragraphs(best).Range
    t=rng.Text.rstrip("\r\n")
    pf=rng.ParagraphFormat
    sz=rng.Font.Size
    # page/section content width
    ps=doc.PageSetup
    cw = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    left = pf.LeftIndent
    print(f"para#{best} len={len(t)} font.size={sz} LeftIndent={left:.1f} colW={cw:.1f} text={t[:30]!r}")
    # line break scan
    base=None;prev=None;breaks=[]
    for k in range(len(t)):
        c=doc.Range(rng.Start+k,rng.Start+k+1)
        ln=c.Information(10)
        if base is None: base=ln
        if prev is not None and ln!=prev: breaks.append(k)
        prev=ln
    nlines=(prev-base+1) if prev is not None else 1
    print(f"display_lines={nlines} break_chars={breaks}")
    if breaks:
        # line-2 width if exists (cleanest, no hanging)
        if len(breaks)>=2:
            l2=breaks[1]-breaks[0]
            cont_w = cw - left
            print(f"line2 chars={l2}  cont_width={cont_w:.1f}  eff_charW={cont_w/l2:.2f}pt")
        l1=breaks[0]
        first_w = cw - (left + pf.FirstLineIndent)
        print(f"line1 chars={l1}  line1_width={first_w:.1f}  eff_charW={first_w/l1:.2f}pt")
    if LP_TW: print(f"linePitch={LP_TW}tw={LP_TW/20:.2f}pt  default_fs guess=10.5  body_fs={sz}")
finally:
    doc.Close(False); word.Quit()
