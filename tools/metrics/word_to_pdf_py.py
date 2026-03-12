"""
Convert test .docx files to PDF using Word COM Automation via pywin32.
Run on Windows with Microsoft Word installed.
"""
import os
import sys
import glob
import time

import win32com.client as win32

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(SCRIPT_DIR, "docx_tests")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output", "pdfs")

WD_EXPORT_FORMAT_PDF = 17
WD_DO_NOT_SAVE_CHANGES = 0


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    docx_files = sorted(glob.glob(os.path.join(INPUT_DIR, "*.docx")))
    if not docx_files:
        print(f"No .docx files found in {INPUT_DIR}")
        sys.exit(1)

    print(f"Starting Microsoft Word...")
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    total = len(docx_files)
    success = 0

    for i, docx_path in enumerate(docx_files, 1):
        filename = os.path.basename(docx_path)
        pdf_name = os.path.splitext(filename)[0] + ".pdf"
        pdf_path = os.path.join(OUTPUT_DIR, pdf_name)

        print(f"[{i}/{total}] {filename} ... ", end="", flush=True)

        try:
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs2(os.path.abspath(pdf_path), FileFormat=WD_EXPORT_FORMAT_PDF)
            doc.Close(WD_DO_NOT_SAVE_CHANGES)
            print("OK")
            success += 1
        except Exception as e:
            print(f"FAILED: {e}")

    word.Quit()
    print(f"\nDone. {success}/{total} PDFs written to {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
