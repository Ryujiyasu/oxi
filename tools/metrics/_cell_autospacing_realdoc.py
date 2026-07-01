# -*- coding: utf-8 -*-
"""Confirm cell-autospacing value on the 2 REAL corpus docs (Ra, 2026-07-01).

For every paragraph whose pPr has before/afterAutospacing="1", report Word's
RESOLVED Format.SpaceBefore/After (the auto value) and whether it sits in a table
cell (Information(12)=wdWithInTable). This confirms the synthetic finding
(cell autospace = body 13.75) on real documents.
"""
import os, sys, io, zipfile, re
import win32com.client
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')

DOCS = [
    'tools/golden-test/documents/docx/29dc6e8943fe_order_01.docx',
    'tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx',
]

def autospacing_para_texts(docx):
    """Return the set of (stripped) paragraph text prefixes whose pPr carries
    a before/afterAutospacing="1" — used to match COM paragraphs."""
    out = []
    with zipfile.ZipFile(docx) as z:
        xml = z.read('word/document.xml').decode('utf-8','replace')
    # crude paragraph split
    for pm in re.finditer(r'<w:p[ >].*?</w:p>', xml, re.S):
        seg = pm.group(0)
        if re.search(r'(before|after)[Aa]utospacing\s*=\s*"(1|true|on)"', seg):
            txt = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', seg, re.S))
            which = []
            if re.search(r'before[Aa]utospacing\s*=\s*"(1|true|on)"', seg): which.append('B')
            if re.search(r'after[Aa]utospacing\s*=\s*"(1|true|on)"', seg): which.append('A')
            out.append((txt.strip(), ''.join(which)))
    return out

def main():
    word = win32com.client.Dispatch("Word.Application"); word.Visible=False; word.DisplayAlerts=False
    try:
        for docx in DOCS:
            print("\n==== %s ====" % os.path.basename(docx))
            want = autospacing_para_texts(docx)
            print("  %d autospacing paragraphs in XML" % len(want))
            doc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            try:
                matched = 0
                for p in doc.Paragraphs:
                    t = p.Range.Text.strip()
                    for (wt, which) in want:
                        if wt and t == wt:
                            try:
                                in_tbl = p.Range.Information(12)  # wdWithInTable
                            except Exception:
                                in_tbl = '?'
                            sb = p.Format.SpaceBefore
                            sa = p.Format.SpaceAfter
                            fs = p.Range.Font.Size
                            print("  [%s] in_cell=%-5s fs=%-5s SpaceBefore=%-7.2f SpaceAfter=%-7.2f  text=%r"
                                  % (which, in_tbl, fs, sb, sa, (t[:24])))
                            matched += 1
                            break
                print("  matched %d/%d" % (matched, len(want)))
            finally:
                doc.Close(False)
    finally:
        word.Quit()

if __name__ == '__main__':
    main()
