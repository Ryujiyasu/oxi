# -*- coding: utf-8 -*-
"""Measure Word per-char wrap for the s562 jc=left yakumono repros: does the
trailing 、 (after K あ) stay on line 1 (oikomi/hang) or wrap (oidashi)?"""
import os,sys,glob
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8')
wd=win32.gencache.EnsureDispatch('Word.Application'); wd.Visible=False
base=os.path.abspath('tools/golden-test/repros/s562_jcleft_yakumono')
print('variant: line1_last_char  comma_line  comma_x  (where 、 lands)')
for f in sorted(glob.glob(base+'/*.docx')):
    doc=wd.Documents.Open(f,ReadOnly=True)
    para=doc.Paragraphs(1).Range
    s=para.Text
    # find the comma index
    ci=s.find('、')
    # per-char line (Information(6)) for chars around the comma
    def yx(i):
        ch=doc.Range(para.Start+i,para.Start+i+1); return ch.Information(6),ch.Information(5)
    # line of char K-1 (last あ before comma) and comma
    y_prev,x_prev=yx(ci-1); y_c,x_c=yx(ci)
    comma_line = 1 if abs(y_c-yx(0)[0])<3 else 2
    # detect: is comma on same line as the あ before it?
    same = abs(y_c-y_prev)<3
    print('%-18s K_chars_before_comma=%d  comma_same_line_as_prev=%s comma_x=%.1f prev_x=%.1f'%(
        os.path.basename(f)[:-11], ci, same, x_c, x_prev))
    doc.Close(False)
wd.Quit()
