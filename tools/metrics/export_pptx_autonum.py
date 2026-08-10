"""Export autonum.pptx to PDF via PowerPoint COM (fresh DispatchEx instance)."""
import os
import sys

import win32com.client


def main(pptx_path, pdf_path):
    pptx_path = os.path.abspath(pptx_path)
    pdf_path = os.path.abspath(pdf_path)
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(pptx_path, WithWindow=False)
        pres.SaveAs(pdf_path, 32)  # ppSaveAsPDF
        pres.Close()
    finally:
        app.Quit()


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
