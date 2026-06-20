# COM-measure aiguideline_komon: the ①②③ lines around "６．禁止事項" (word_i=53).
# Per-paragraph: page, vertical Y, line height (Y gap to next), font name, font.size, text head.
import sys, win32com.client as w
DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\aiguideline_komon.docx"
app = w.Dispatch("Word.Application"); app.Visible = False
try:
    doc = app.Documents.Open(DOC, ReadOnly=True)
    paras = doc.Paragraphs
    n = paras.Count
    print(f"n_paras={n}")
    def yof(rng):
        return doc.Range(rng.Start, rng.Start).Information(6)  # vertical pos rel page
    def pof(rng):
        return doc.Range(rng.Start, rng.Start).Information(3)  # page number
    # find paragraphs whose text starts with a circled number or 禁止
    rows=[]
    prev=None
    for i in range(1, n+1):
        p = paras(i)
        txt = (p.Range.Text or "").replace("\r","").replace("\x07","")
        head = txt[:24]
        rng = p.Range
        y = yof(rng); pg = pof(rng)
        # font of first char
        fc = doc.Range(rng.Start, min(rng.Start+1, rng.End))
        try:
            fn = fc.Font.NameFarEast or fc.Font.Name
            fnasc = fc.Font.Name
            fsz = fc.Font.Size
        except Exception:
            fn="?"; fnasc="?"; fsz=0
        rows.append((i,pg,round(y,2),head,fn,fnasc,fsz))
    # print rows near "禁止事項" and any with circled numbers
    circ = set("①②③④⑤⑥⑦⑧⑨⑩")
    for idx,(i,pg,y,head,fn,fnasc,fsz) in enumerate(rows):
        flag = ("禁止" in head) or any(c in head for c in circ)
        ctx = flag or (idx>0 and (("禁止" in rows[idx-1][3]) or any(c in rows[idx-1][3] for c in circ)))
        if flag or ctx:
            gap = (rows[idx+1][2]-y) if idx+1<len(rows) and rows[idx+1][1]==pg else None
            gaps = f"{gap:.2f}" if gap is not None else "   -"
            print(f"  i={i:3d} p{pg} y={y:7.2f} gap={gaps} sz={fsz} fe={fn!r} asc={fnasc!r} :: {head!r}")
    doc.Close(False)
finally:
    app.Quit()
