# -*- coding: utf-8 -*-
"""Confirm via RENDERED Y-GAPS (not Format.SpaceBefore) what Word actually renders
around the autospacing cell paragraphs in the 2 real corpus docs. (Ra 2026-07-01)

Identifies the autospacing paragraphs by document order (incl. empties) and prints
the Information(6) Y-gap to the previous & next paragraph, plus their text & whether
in a table.
"""
import os, sys, io, zipfile, re
import win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')

DOCS = [
    'tools/golden-test/documents/docx/29dc6e8943fe_order_01.docx',
    'tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx',
]

def autospacing_indices(docx):
    """0-based paragraph indices (document order, incl empties) whose pPr has autospacing."""
    with zipfile.ZipFile(docx) as z:
        xml = z.read('word/document.xml').decode('utf-8','replace')
    # iterate body paragraphs in order (only those directly inside w:body OR cells — all <w:p>)
    idxs = []
    for i, pm in enumerate(re.finditer(r'<w:p(?:\s[^>]*)?>.*?</w:p>', xml, re.S)):
        seg = pm.group(0)
        if re.search(r'(before|after)[Aa]utospacing\s*=\s*"(1|true|on)"', seg):
            which = ''
            if re.search(r'before[Aa]utospacing\s*=\s*"(1|true|on)"', seg): which+='B'
            if re.search(r'after[Aa]utospacing\s*=\s*"(1|true|on)"', seg): which+='A'
            idxs.append((i, which))
    return idxs

def main():
    word = win32com.client.Dispatch("Word.Application"); word.Visible=False; word.DisplayAlerts=False
    try:
        for docx in DOCS:
            print("\n==== %s ====" % os.path.basename(docx))
            asx = dict(autospacing_indices(docx))
            doc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            try:
                paras = list(doc.Paragraphs)
                Y = []
                for p in paras:
                    sr = doc.Range(p.Range.Start, p.Range.Start)
                    Y.append(sr.Information(6))
                for i, p in enumerate(paras):
                    if i not in asx:
                        continue
                    t = p.Range.Text.strip()
                    try: in_tbl = p.Range.Information(12)
                    except Exception: in_tbl='?'
                    fs = p.Range.Font.Size
                    prevg = (Y[i]-Y[i-1]) if i>0 else None
                    nextg = (Y[i+1]-Y[i]) if i+1<len(Y) else None
                    pt = paras[i-1].Range.Text.strip()[:10] if i>0 else ''
                    nt = paras[i+1].Range.Text.strip()[:10] if i+1<len(paras) else ''
                    pg = '%.2f'%prevg if prevg is not None else '-'
                    ng = '%.2f'%nextg if nextg is not None else '-'
                    print("  idx%-4d [%s] cell=%-5s fs=%-5s  prevGap=%-7s nextGap=%-7s  self=%r prev=%r next=%r"
                          % (i, asx[i], in_tbl, fs, pg, ng, t[:16], pt, nt))
            finally:
                doc.Close(False)
    finally:
        word.Quit()

if __name__ == '__main__':
    main()
