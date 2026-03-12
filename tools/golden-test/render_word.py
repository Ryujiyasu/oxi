#!/usr/bin/env python3
"""Render all golden test docx files using Microsoft Word COM API.

Converts each docx to PDF via Word, then renders page 1 to PNG at 150 DPI
using pdf2image (poppler) or PyMuPDF.
"""
import os
import sys
import time
from pathlib import Path

def render_with_word(docx_files, output_dir):
    """Use Word COM to save each docx as PDF, then convert to PNG."""
    import win32com.client
    import fitz  # PyMuPDF

    output_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir = output_dir.parent / "word_pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone

    total = len(docx_files)
    for i, docx_path in enumerate(docx_files):
        stem = docx_path.stem
        pdf_path = pdf_dir / f"{stem}.pdf"
        png_path = output_dir / f"{stem}.png"

        if png_path.exists():
            print(f"  [{i+1}/{total}] SKIP {stem[:40]} (exists)")
            continue

        print(f"  [{i+1}/{total}] {stem[:50]}...", end=" ", flush=True)

        try:
            doc = word.Documents.Open(str(docx_path.resolve()))
            # Save as PDF (wdFormatPDF = 17)
            doc.SaveAs2(str(pdf_path.resolve()), FileFormat=17)
            doc.Close(SaveChanges=0)
            print("PDF", end=" ", flush=True)

            # Convert first page of PDF to PNG at 150 DPI
            pdf_doc = fitz.open(str(pdf_path))
            page = pdf_doc[0]
            # 150 DPI = 150/72 zoom factor
            zoom = 150 / 72
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            pix.save(str(png_path))
            pdf_doc.close()
            print(f"PNG ({pix.width}x{pix.height})")

        except Exception as e:
            print(f"ERROR: {e}")

    word.Quit()
    print(f"\nDone. {total} files processed.")


def main():
    script_dir = Path(__file__).resolve().parent
    output_dir = script_dir / "pixel_output" / "word"
    docx_dir = script_dir / "documents" / "docx"
    fixtures_dir = script_dir.parent.parent / "tests" / "fixtures"

    # Collect all docx files that have LO renders (to compare same set)
    lo_dir = script_dir / "pixel_output" / "libreoffice"

    test_files = []
    for f in sorted(fixtures_dir.glob("*.docx")):
        if lo_dir.exists() and (lo_dir / f"{f.stem}.png").exists():
            test_files.append(f)
        elif not lo_dir.exists():
            test_files.append(f)

    if docx_dir.exists():
        for f in sorted(docx_dir.glob("*.docx")):
            if lo_dir.exists() and (lo_dir / f"{f.stem}.png").exists():
                test_files.append(f)
            elif not lo_dir.exists():
                test_files.append(f)

    if not test_files:
        print("No docx files found!")
        sys.exit(1)

    print(f"Rendering {len(test_files)} files with Word...")
    render_with_word(test_files, output_dir)


if __name__ == "__main__":
    main()
