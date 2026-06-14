# -*- coding: utf-8 -*-
"""Generic: export a doc to PDF via Word, return per-page body lines (text)."""
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
DOCX = os.path.abspath(sys.argv[1])
PDF  = os.path.abspath(sys.argv[2])
if not os.path.exists(PDF) or '--reexport' in sys.argv:
    import win32com.client as win32
    w = win32.DispatchEx('Word.Application'); w.Visible=False
    try:
        d = w.Documents.Open(DOCX, ReadOnly=True)
        d.ExportAsFixedFormat(PDF, 17)
        d.Close(False)
    finally:
        w.Quit()
    print('exported', PDF)
