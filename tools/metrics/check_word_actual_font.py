"""Verify what font Word ACTUALLY uses for CJK runs in regressed 177-doc files.

Word's COM Range.Font.Name reports the font Word resolved for that text,
including theme/style/docDefaults inheritance.
"""
import win32com.client
import time
import sys

sys.stdout.reconfigure(encoding='utf-8')

DOCS = [
    'tools/golden-test/documents/docx/04b88e7e0b25_index-19.docx',
    'tools/golden-test/documents/docx/4a36b62555f2_kyodokenkyuyoushiki10.docx',
]

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

import os
for d in DOCS:
    abs_d = os.path.abspath(d)
    print('\n========', os.path.basename(d))
    doc = word.Documents.Open(abs_d, ReadOnly=True)
    time.sleep(0.5)

    # Sample first few paragraphs with CJK content
    for pi in range(1, min(6, doc.Paragraphs.Count + 1)):
        p = doc.Paragraphs(pi)
        text = p.Range.Text.strip()
        if not text or len(text) < 2:
            continue
        # Pick first CJK char
        cjk_idx = None
        for i, ch in enumerate(text):
            if '\u3000' <= ch <= '\uffef' or '\u4e00' <= ch <= '\u9fff':
                cjk_idx = i
                break
        if cjk_idx is None:
            continue
        # Get Range for that single char
        cr = p.Range.Duplicate
        cr.Start = p.Range.Start + cjk_idx
        cr.End = cr.Start + 1
        try:
            name_ascii = cr.Font.Name
            name_ea = cr.Font.NameFarEast
        except Exception as e:
            name_ascii = name_ea = f'(err {e})'
        ch = text[cjk_idx]
        print(f'  P{pi} char {cjk_idx} "{ch}": Font.Name={name_ascii!r}, NameFarEast={name_ea!r}')

    # Also check what docDefaults says via Word's Style.Font
    try:
        normal = doc.Styles('Normal')
        print(f'  Normal style: Font.Name={normal.Font.Name!r}, NameFarEast={normal.Font.NameFarEast!r}')
    except Exception as e:
        print(f'  Normal style err: {e}')

    doc.Close(SaveChanges=False)

word.Quit()
