import os, win32com.client as win32
DOCX=os.path.abspath("tools/golden-test/documents/docx/harassmanual_001466344.docx")
w=win32.gencache.EnsureDispatch("Word.Application");w.Visible=False
d=w.Documents.Open(DOCX,ReadOnly=True)
try:
    tgt=None
    for i in range(1,d.Paragraphs.Count+1):
        if d.Paragraphs(i).Range.Text.startswith("相談者の直面する"): tgt=i;break
    r=d.Paragraphs(tgt).Range; t=r.Text.rstrip("\r\n")
    print("para16 len=%d text=%r"%(len(t),t[:40]))
    # per-char advance for line 1 (first ~34 chars)
    prev=None
    for k in range(0,34):
        x=d.Range(r.Start+k,r.Start+k+1).Information(5)
        adv = (x-prev) if prev is not None else 0
        if prev is not None:
            mark = "  <<< COMPRESSED" if (t[k-1] in "、。，．" and adv<9) else ""
            print("  %r adv=%.2f%s"%(t[k-1],adv,mark))
        prev=x
    # also where does line 1 end (line number per char)
    print("--- line numbers ---")
    base=None;pl=None
    for k in range(len(t)):
        ln=d.Range(r.Start+k,r.Start+k+1).Information(10)
        if base is None: base=ln
        if pl is not None and ln!=pl: print("  break before char %d %r"%(k,t[k]))
        pl=ln
finally:
    d.Close(False);w.Quit()
