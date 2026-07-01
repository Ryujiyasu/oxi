# -*- coding: utf-8 -*-
"""Cell autospacing follow-up (Ra, 2026-07-01): nail the 8pt outlier, the
direct-vs-style gating (S675 gate-safety), and font-size attribution.
"""
import os, sys, io, zipfile
sys.path.insert(0, 'tools/metrics')
import win32com.client
from mixedh_lineplace import build_generic, OUT, WNS, MNS
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')

FONT = 'ＭＳ 明朝'
def rpr(sz): return ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/>'%(FONT,FONT,FONT,sz))
def para(text, sz=22, sp='', pstyle=''):
    spx = ('<w:spacing %s/>'%sp) if sp else ''
    ps = ('<w:pStyle w:val="%s"/>'%pstyle) if pstyle else ''
    return ('<w:p><w:pPr>%s%s<w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'%(ps,spx,rpr(sz),rpr(sz),text))
def cell(inner):
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>'%inner)
def ys(word, docx):
    doc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True); out=[]
    try:
        for p in doc.Paragraphs:
            sr = doc.Range(p.Range.Start, p.Range.Start)
            out.append((sr.Information(6), p.Range.Text.strip()))
    finally: doc.Close(False)
    return out

# styles.xml with a custom paragraph style "WebStyle" carrying after/beforeAutospacing,
# plus a docDefaults sz, to test (a) style-inherited autospacing, (b) docDefaults effect.
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '</Relationships>')

def styles_xml(default_sz=None, web_attrs=''):
    dd_sz = ('<w:sz w:val="%d"/>'%default_sz) if default_sz else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles %s>'
            '<w:docDefaults><w:rPrDefault><w:rPr>%s</w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="WebStyle"><w:name w:val="WebStyle"/>'
            '<w:pPr><w:spacing %s/></w:pPr></w:style>'
            '</w:styles>'%(WNS, dd_sz, web_attrs))

def build_styled(name, body_xml, default_sz=None, web_attrs=''):
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document %s %s><w:body>%s%s</w:body></w:document>'%(WNS,MNS,body_xml,sect))
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOCRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', styles_xml(default_sz, web_attrs))
    return p

def main():
    word = win32com.client.Dispatch("Word.Application"); word.Visible=False; word.DisplayAlerts=False
    try:
        # ---- (A) fine small-size sweep, ALL THREE paras same size (no cross-size) ----
        print("=== (A) cell afterAuto, TOP/MID/BOT ALL same size ===")
        print("  %-6s %-9s %-9s %-9s" % ("sz","auto","norm","afterAuto"))
        for szhalf in (12,14,16,18,20,21,22,24):
            a = ys(word, build_generic('cas2_a_%d.docx'%szhalf,
                cell(para('Ｔ',szhalf)+para('Ｍ',szhalf,'w:afterAutospacing="1"')+para('Ｂ',szhalf))))
            n = ys(word, build_generic('cas2_n_%d.docx'%szhalf,
                cell(para('Ｔ',szhalf)+para('Ｍ',szhalf)+para('Ｂ',szhalf))))
            gA=a[2][0]-a[1][0]; gN=n[2][0]-n[1][0]
            print("  %-6.1f %-9.2f %-9.2f %-9.2f" % (szhalf/2.0,gA,gN,gA-gN))

        # ---- (B) docDefaults sz effect on the auto value (para itself 11pt) ----
        print("\n=== (B) docDefaults sz sweep (MID para fixed 11pt, afterAuto) ===")
        print("  %-10s %-9s %-9s %-9s" % ("ddSz(half)","auto","norm","afterAuto"))
        for dd in (16, 20, 22, 24, 28):
            a = ys(word, build_styled('cas2_dd_a_%d.docx'%dd,
                cell(para('Ｔ',22)+para('Ｍ',22,'w:afterAutospacing="1"')+para('Ｂ',22)), default_sz=dd))
            n = ys(word, build_styled('cas2_dd_n_%d.docx'%dd,
                cell(para('Ｔ',22)+para('Ｍ',22)+para('Ｂ',22)), default_sz=dd))
            gA=a[2][0]-a[1][0]; gN=n[2][0]-n[1][0]
            print("  %-10d %-9.2f %-9.2f %-9.2f" % (dd,gA,gN,gA-gN))

        # ---- (C) STYLE-inherited autospacing in a cell (gate safety) ----
        print("\n=== (C) STYLE-inherited (WebStyle) autospacing in a cell ===")
        # MID uses pStyle=WebStyle (after+before auto). Does Word apply it?
        a = ys(word, build_styled('cas2_style.docx',
            cell(para('Ｔ',22)+para('Ｍ',22,pstyle='WebStyle')+para('Ｂ',22)),
            web_attrs='w:beforeAutospacing="1" w:afterAutospacing="1"'))
        n = ys(word, build_styled('cas2_style_n.docx',
            cell(para('Ｔ',22)+para('Ｍ',22)+para('Ｂ',22)),
            web_attrs='w:beforeAutospacing="1" w:afterAutospacing="1"'))
        gTM_a=a[1][0]-a[0][0]; gMB_a=a[2][0]-a[1][0]
        gTM_n=n[1][0]-n[0][0]; gMB_n=n[2][0]-n[1][0]
        print("  TOP->MID  style=%.2f norm=%.2f  before=%.2f" % (gTM_a,gTM_n,gTM_a-gTM_n))
        print("  MID->BOT  style=%.2f norm=%.2f  after=%.2f"  % (gMB_a,gMB_n,gMB_a-gMB_n))

        # ---- (D) DIRECT autospacing in cell WITH a docDefaults+styles doc (parity) ----
        print("\n=== (D) DIRECT autospacing in a cell (styled doc, parity check) ===")
        a = ys(word, build_styled('cas2_direct.docx',
            cell(para('Ｔ',22)+para('Ｍ',22,'w:afterAutospacing="1"')+para('Ｂ',22))))
        n = ys(word, build_styled('cas2_direct_n.docx',
            cell(para('Ｔ',22)+para('Ｍ',22)+para('Ｂ',22))))
        print("  MID->BOT  direct=%.2f norm=%.2f  after=%.2f"
              % (a[2][0]-a[1][0], n[2][0]-n[1][0], (a[2][0]-a[1][0])-(n[2][0]-n[1][0])))
    finally:
        word.Quit()

if __name__ == '__main__':
    main()
