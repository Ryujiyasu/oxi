# -*- coding: utf-8 -*-
"""THE dominant corpus case: a SOLE empty autospacing para in a cell (29dc6e: 4/6).
Measure the cell's contributed height (Y below table - Y above table) to determine
whether Word applies / suppresses before & after for a sole (first==last) cell para,
vs explicit-only and plain. (Ra 2026-07-01)
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
def pcontent(text,spx): return '<w:p><w:pPr><w:spacing %s/><w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr><w:t>%s</w:t></w:r></w:p>'%(spx,rpr(),rpr(),text)
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

CORPUS='w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="exact"'
PLAIN='w:line="240" w:lineRule="exact"'
EXPLICIT='w:before="100" w:after="100" w:line="240" w:lineRule="exact"'
BAONLY='w:beforeAutospacing="1" w:afterAutospacing="1" w:line="240" w:lineRule="exact"'

def cellh(word, name, inner):
    """table contributed height = Y(below) - Y(above) - above_line_height."""
    y = ys(word, build_generic(name, pn('ＡＢＯＶＥ')+inner+pn('ＢＥＬＯＷ')))
    # y[0]=ABOVE, y[1]=the (sole) cell para, y[2]=BELOW  (cell para is its own paragraph)
    above_to_below = y[-1][0]-y[0][0]
    above_to_cell = y[1][0]-y[0][0]
    cell_to_below = y[-1][0]-y[1][0]
    return above_to_below, above_to_cell, cell_to_below

def main():
    word=win32com.client.Dispatch("Word.Application"); word.Visible=False; word.DisplayAlerts=False
    try:
        print("SOLE empty cell para — table contributed geometry (ABOVE/BELOW are 11pt body):")
        print("  %-26s %-12s %-12s %-12s" % ("pattern","above->below","above->cell","cell->below"))
        for label, spx in [('corpus(b100ba+a100aa)',CORPUS),('ba_only',BAONLY),('explicit(b100+a100)',EXPLICIT),('plain',PLAIN)]:
            ab,ac,cb = cellh(word, 'csole_%s.docx'%label[:8], cell(pempty(spx)))
            print("  %-26s %-12.2f %-12.2f %-12.2f" % (label, ab, ac, cb))
        # also: SOLE NON-EMPTY content cell para (the 29dc6e idx94-ish if it were sole)
        print("\nSOLE non-empty cell para:")
        for label, spx in [('corpus',CORPUS),('explicit',EXPLICIT),('plain',PLAIN)]:
            ab,ac,cb = cellh(word, 'csolec_%s.docx'%label[:6], cell(pcontent('文',spx)))
            print("  %-26s above->below=%-8.2f" % (label, ab))
        # FIRST of 2 (before suppressed, after interior): cell[ contentPara(auto), empty ]
        print("\nFIRST-of-2 (29dc6e idx 1/2 case): cell[ para(b100ba+a100aa), Ｂ ]")
        for label, spx in [('corpus',CORPUS),('explicit',EXPLICIT),('plain',PLAIN)]:
            y = ys(word, build_generic('cf2_%s.docx'%label[:6], pn('上')+cell(pcontent('先',spx)+pn('後'))+pn('下')))
            # y: 上, 先, 後, 下
            print("  %-12s 上->先(before)=%-8.2f 先->後(after)=%-8.2f" % (label, y[1][0]-y[0][0], y[2][0]-y[1][0]))
    finally:
        word.Quit()
if __name__=='__main__': main()
