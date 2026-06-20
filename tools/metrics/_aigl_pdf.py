# Export aiguideline to PDF via Word COM, then use fitz to find the ① glyph's
# render font + bbox, and compare line heights of ①-lines vs plain body lines.
import os, win32com.client as w, fitz
DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\aiguideline_komon.docx"
PDF = r"C:\tmp\aigl.pdf"
os.makedirs(r"C:\tmp", exist_ok=True)
app = w.Dispatch("Word.Application"); app.Visible=False
try:
    d=app.Documents.Open(DOC, ReadOnly=True)
    d.ExportAsFixedFormat(PDF, 17)  # wdExportFormatPDF
    d.Close(False)
finally:
    app.Quit()
doc=fitz.open(PDF)
print("pages:", doc.page_count)
circ=set("①②③④⑤⑥⑦⑧⑨⑩")
for pno in range(min(3,doc.page_count)):
    pg=doc[pno]
    dd=pg.get_text("dict")
    print(f"--- page {pno+1} ---")
    for b in dd["blocks"]:
        if b.get("type")!=0: continue
        for ln in b["lines"]:
            spans=ln["spans"]
            txt="".join(s["text"] for s in spans)
            has_circ=any(c in txt for c in circ)
            if not has_circ: continue
            for s in spans:
                if any(c in s["text"] for c in circ):
                    bb=s["bbox"]
                    print(f"  '{s['text'][:14]}' font={s['font']!r} size={s['size']:.2f} y0={bb[1]:.2f} y1={bb[3]:.2f} h={bb[3]-bb[1]:.2f} asc={s.get('ascender')} desc={s.get('descender')}")
                    break
