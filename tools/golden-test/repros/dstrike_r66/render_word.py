"""Render a single .docx via Word ExportAsFixedFormat → PDF → PNG.

Mirrors pipeline/word_renderer.py path for one file, to a known location
so we can visually compare with Oxi GDI/DWrite renders.
"""
import os
import sys
import time
import subprocess
from pathlib import Path

import win32com.client
import pythoncom
import fitz


DPI = 144


def kill_word():
    subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
    time.sleep(1.0)


def render(docx_path: str, out_dir: str):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    docx_path = os.path.abspath(docx_path)

    kill_word()
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    word.AutomationSecurity = 3
    word.Options.UpdateLinksAtOpen = False
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True,
                                  AddToRecentFiles=False, ConfirmConversions=False)
        try:
            page_count = int(doc.ComputeStatistics(2))
            print(f"  {page_count} pages")
            for page_num in range(1, page_count + 1):
                pdf_path = out_dir / f"word_p{page_num}.pdf"
                png_path = out_dir / f"word_p{page_num}.png"
                doc.ExportAsFixedFormat(
                    OutputFileName=str(pdf_path),
                    ExportFormat=17, OpenAfterExport=False,
                    OptimizeFor=0, Range=3, From=page_num, To=page_num,
                )
                d = fitz.open(str(pdf_path))
                zoom = DPI / 72
                pix = d[0].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                pix.save(str(png_path))
                d.close()
                pdf_path.unlink()
                print(f"  saved {png_path}")
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    docx = sys.argv[1] if len(sys.argv) > 1 else "DS_full.docx"
    out = sys.argv[2] if len(sys.argv) > 2 else r"C:\Users\ryuji\AppData\Local\Temp\dstrike_render"
    render(docx, out)
