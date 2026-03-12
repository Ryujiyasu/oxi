#!/usr/bin/env python3
"""Minimal COM test to verify Word automation works."""
import os, sys, time, tempfile
import pythoncom
pythoncom.CoInitialize()
import win32com.client
from docx import Document

# Create minimal docx
doc = Document()
doc.add_paragraph("Hello World")
tmp_docx = os.path.join(tempfile.gettempdir(), "test_minimal.docx")
tmp_pdf = os.path.join(tempfile.gettempdir(), "test_minimal.pdf")
doc.save(tmp_docx)

print(f"docx: {tmp_docx}")
print(f"pdf:  {tmp_pdf}")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

abs_docx = os.path.abspath(tmp_docx).replace("/", "\\")
abs_pdf = os.path.abspath(tmp_pdf).replace("/", "\\")

print("Opening document...")
wdoc = word.Documents.Open(abs_docx)
print(f"  doc type: {type(wdoc)}")
time.sleep(2)

print("Saving as PDF...")
wdoc.SaveAs(abs_pdf, FileFormat=17)
time.sleep(1)

print("Closing...")
wdoc.Close(False)
word.Quit()

print(f"PDF exists: {os.path.exists(tmp_pdf)}")
if os.path.exists(tmp_pdf):
    print(f"PDF size: {os.path.getsize(tmp_pdf)} bytes")
    import fitz
    pdoc = fitz.open(tmp_pdf)
    print(f"Pages: {len(pdoc)}")
    page = pdoc[0]
    blocks = page.get_text("dict")["blocks"]
    for b in blocks:
        if "lines" in b:
            for line in b["lines"]:
                text = "".join(s["text"] for s in line["spans"])
                y = line["bbox"][1]
                print(f"  y={y:.2f} text=\"{text}\"")
    pdoc.close()

os.unlink(tmp_docx)
os.unlink(tmp_pdf)
print("SUCCESS")
