# -*- coding: utf-8 -*-
"""Measure real-doc consecutive-autospacing gaps (harassbosi/b837) in Word."""
import os, sys, io
import win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')

def measure(word, path, nmax=14):
    doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
    rows = []
    try:
        ps = doc.Paragraphs
        prev_y = None; prev_pg = None
        for i in range(1, min(ps.Count, nmax)+1):
            p = ps(i)
            r = p.Range
            sr = doc.Range(r.Start, r.Start)
            y = sr.Information(6)   # vertical pos rel page
            pg = sr.Information(3)  # page number
            sid = p.Style.NameLocal
            txt = r.Text.strip()[:14]
            gap = (y - prev_y) if (prev_y is not None and pg == prev_pg) else None
            rows.append((i, pg, round(y,2), (round(gap,2) if gap is not None else '-PB-'), sid, txt))
            prev_y, prev_pg = y, pg
    finally:
        doc.Close(False)
    return rows

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False; word.DisplayAlerts = False
    try:
        for path in ['tools/golden-test/documents/docx/harassbosi_002140020.docx',
                     'tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx']:
            print('===', os.path.basename(path), '===')
            print('%-3s %-4s %-8s %-8s %-16s %s' % ('i','pg','y','gap','style','text'))
            for r in measure(word, path):
                print('%-3d %-4s %-8s %-8s %-16s %s' % r)
            print()
    finally:
        word.Quit()

if __name__ == '__main__':
    main()
