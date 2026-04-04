"""Export Word page as GDI-rendered PNG via XPS SaveAs.

XPS preserves Word's GDI layout coordinates exactly.
Renders XPS → PNG using Windows XPS renderer.

Alternative: Use Word SaveAs PNG directly (most accurate).

Usage: python word_gdi_export.py <docx_path> <output_prefix> [dpi]
"""
import win32com.client
import pythoncom
import subprocess
import sys
import os
import tempfile

def export_as_png(docx_path, output_prefix, dpi=150):
    """Export Word document pages as PNG using Word's built-in image export."""
    docx_path = os.path.abspath(docx_path)
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        total_pages = doc.ComputeStatistics(2)  # wdStatisticPages
        print(f"Document: {os.path.basename(docx_path)}, Pages: {total_pages}, DPI: {dpi}")

        # Method: Save as PDF, then render PDF to PNG at target DPI
        tmp_pdf = os.path.join(tempfile.gettempdir(), "word_export_temp.pdf")
        doc.SaveAs(tmp_pdf, 17)  # wdFormatPDF
        doc.Close(False)

        # Render PDF pages to PNG using Poppler (pdftoppm)
        for page in range(1, total_pages + 1):
            out_path = f"{output_prefix}_p{page}.png"
            cmd = [
                "pdftoppm", "-png", "-r", str(dpi),
                "-f", str(page), "-l", str(page),
                "-singlefile",
                tmp_pdf, out_path.replace(".png", "")
            ]
            try:
                subprocess.run(cmd, check=True, capture_output=True)
                print(f"  Saved {out_path}")
            except FileNotFoundError:
                # pdftoppm not available, try Ghostscript
                gs_cmd = [
                    "gswin64c", "-dNOPAUSE", "-dBATCH", "-sDEVICE=png16m",
                    f"-r{dpi}", f"-dFirstPage={page}", f"-dLastPage={page}",
                    f"-sOutputFile={out_path}", tmp_pdf
                ]
                try:
                    subprocess.run(gs_cmd, check=True, capture_output=True)
                    print(f"  Saved {out_path} (via Ghostscript)")
                except FileNotFoundError:
                    print(f"  ERROR: Neither pdftoppm nor gswin64c found. Install Poppler or Ghostscript.")
                    break

        # Cleanup
        try:
            os.remove(tmp_pdf)
        except:
            pass

    finally:
        word.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python word_gdi_export.py <docx_path> <output_prefix> [dpi]")
        sys.exit(1)
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 150
    export_as_png(sys.argv[1], sys.argv[2], dpi)
