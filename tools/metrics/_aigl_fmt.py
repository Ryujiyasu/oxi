import win32com.client as w
DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\aiguideline_komon.docx"
app=w.Dispatch("Word.Application"); app.Visible=False
try:
    d=app.Documents.Open(DOC, ReadOnly=True)
    P=d.Paragraphs
    for i in [11,25,26,27,53,54,59,60]:
        p=P(i); pf=p.Format; rng=p.Range
        txt=(rng.Text or "").replace("\r","").replace("\x07","")[:18]
        sb=pf.SpaceBefore; sa=pf.SpaceAfter
        lsr=pf.LineSpacingRule; ls=pf.LineSpacing
        try: style=p.Style.NameLocal
        except: style="?"
        # first-char font detail
        fc=d.Range(rng.Start, min(rng.Start+1,rng.End))
        bold = fc.Font.Bold
        print(f"i={i:3d} sb={sb:5.2f} sa={sa:5.2f} lsRule={lsr} ls={ls:6.2f} bold={bold} style={style!r} :: {txt!r}")
    # circled-number paragraph contextualSpacing?
    print("--- raw spacing XML probe (i=25,27) ---")
    for i in [25,27,11]:
        p=P(i)
        xml=p.Range.ParagraphFormat
        # use w:contextualSpacing presence via WordOpenXML on the paragraph
        ox=p.Range.WordOpenXML
        import re
        cs = 'contextualSpacing' in ox
        sbm = re.search(r'w:spacing[^>]*', ox)
        print(f"  i={i} contextualSpacing={cs} spacing_tag={sbm.group(0)[:120] if sbm else None}")
    d.Close(False)
finally:
    app.Quit()
