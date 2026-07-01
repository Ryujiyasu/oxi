# -*- coding: utf-8 -*-
"""The corpus shape: EMPTY autospacing spacer para inside a cell (29dc6e/d4d126 pattern:
before=100 beforeAutospacing after=100 afterAutospacing line=240 lineRule=exact).
Measure the rendered Y-gap an empty spacer adds vs a plain empty spacer. (Ra 2026-07-01)
"""
import os, sys, io, zipfile
import win32com.client
sys.path.insert(0, 'tools/metrics')
from mixedh_lineplace import build_generic
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')

FONT='ＭＳ 明朝'
def rpr(sz=21): return '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/>'%(FONT,FONT,FONT,sz)
def pn(text): return '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr><w:t>%s</w:t></w:r></w:p>'%(rpr(),rpr(),text)
def pempty(spx): return '<w:p><w:pPr><w:spacing %s/><w:rPr>%s</w:rPr></w:pPr></w:p>'%(spx,rpr())
def cell(inner): return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
 '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
 '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr>'
 '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>'%inner)
def ys(word,docx):
    doc=word.Documents.Open(os.path.abspath(docx),ReadOnly=True); out=[]
    try:
        for p in doc.Paragraphs:
            sr=doc.Range(p.Range.Start,p.Range.Start); out.append((sr.Information(6),p.Range.Text.strip()))
    finally: doc.Close(False)
    return out

CORPUS = 'w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="exact"'
PLAIN  = 'w:line="240" w:lineRule="exact"'
EXPLICIT = 'w:before="100" w:after="100" w:line="240" w:lineRule="exact"'

def main():
    word=win32com.client.Dispatch("Word.Application"); word.Visible=False; word.DisplayAlerts=False
    try:
        # ctx body and cell; spacer = empty para between T and B
        for ctx, mk in [('body', lambda b:b), ('cell', lambda b:cell(b))]:
            print("=== %s: empty spacer T / [empty] / B ===" % ctx)
            for label, spx in [('corpus(b100ba+a100aa)',CORPUS),('plain',PLAIN),('explicit(b100+a100)',EXPLICIT)]:
                y = ys(word, build_generic('ces_%s_%s.docx'%(ctx,label[:6]), mk(pn('上')+pempty(spx)+pn('下'))))
                # y[0]=T, y[1]=empty, y[2]=B
                gTE = y[1][0]-y[0][0]   # T -> empty
                gEB = y[2][0]-y[1][0]   # empty -> B
                gTB = y[2][0]-y[0][0]   # T -> B (total the spacer occupies + T line)
                print("  %-22s  T->empty=%-7.2f empty->B=%-7.2f  T->B=%-7.2f" % (label, gTE, gEB, gTB))
            # control: NO spacer para (T directly above B)
            y0 = ys(word, build_generic('ces_%s_nospacer.docx'%ctx, mk(pn('上')+pn('下'))))
            print("  %-22s  T->B(no spacer)=%.2f" % ('(no spacer)', y0[1][0]-y0[0][0]))
            print()
    finally:
        word.Quit()
if __name__=='__main__': main()
