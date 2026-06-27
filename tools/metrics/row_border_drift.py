# -*- coding: utf-8 -*-
import sys, subprocess, os, json, fitz
sys.stdout.reconfigure(encoding="utf-8")
DOCX=sys.argv[1]
# Word PDF
PDF=r"C:\tmp\7ead_word.pdf"
import win32com.client as w
app=w.Dispatch("Word.Application"); app.Visible=False
d=app.Documents.Open(DOCX, ReadOnly=True)
d.ExportAsFixedFormat(PDF, 17)
d.Close(False); app.Quit()
doc=fitz.open(PDF); pg=doc[0]
hl=set()
for dr in pg.get_drawings():
    for it in dr["items"]:
        if it[0]=="l":
            p1,p2=it[1],it[2]
            if abs(p1.y-p2.y)<0.5 and abs(p2.x-p1.x)>20: hl.add(round(p1.y,1))
        elif it[0]=="re":
            r=it[1]
            for y in (r.y0,r.y1):
                if r.width>20: hl.add(round(y,1))
print("WORD horizontal lines (pt):", sorted(hl))
