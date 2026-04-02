"""Export Word page as EMF via CopyAsPicture → clipboard.

Usage: python word_emf_export.py <docx_path> <output_prefix>
Produces: output_prefix_p1.emf, output_prefix_p2.emf, ...
"""
import win32com.client
import win32clipboard
import ctypes
import struct
import sys
import os
import pythoncom
import time

CF_ENHMETAFILE = 14

def export_emf(docx_path, output_prefix):
    docx_path = os.path.abspath(docx_path)
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = True  # Need visible for CopyAsPicture
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        total_pages = doc.ComputeStatistics(2)
        print(f"Pages: {total_pages}")

        for page_num in range(1, total_pages + 1):
            # Navigate to page
            rng = doc.GoTo(1, 2, page_num)  # wdGoToPage, wdGoToAbsolute
            if page_num < total_pages:
                next_page = doc.GoTo(1, 2, page_num + 1)
                rng.End = next_page.Start
            else:
                rng.End = doc.Content.End

            # CopyAsPicture
            rng.CopyAsPicture()
            time.sleep(0.5)

            # Get EMF from clipboard
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(CF_ENHMETAFILE):
                    emf_data = win32clipboard.GetClipboardData(CF_ENHMETAFILE)
                    if isinstance(emf_data, bytes):
                        emf_path = f"{output_prefix}_p{page_num}.emf"
                        with open(emf_path, 'wb') as f:
                            f.write(emf_data)
                        print(f"  Saved {emf_path} ({len(emf_data)} bytes)")
                    else:
                        # Handle case: emf_data is a handle (integer)
                        handle = int(emf_data)
                        gdi32 = ctypes.windll.gdi32
                        gdi32.GetEnhMetaFileBits.restype = ctypes.c_uint
                        gdi32.GetEnhMetaFileBits.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p]
                        size = gdi32.GetEnhMetaFileBits(handle, 0, None)
                        buf = ctypes.create_string_buffer(size)
                        gdi32.GetEnhMetaFileBits(handle, size, buf)
                        emf_path = f"{output_prefix}_p{page_num}.emf"
                        with open(emf_path, 'wb') as f:
                            f.write(buf.raw)
                        print(f"  Saved {emf_path} ({size} bytes)")
                else:
                    avail = []
                    fmt = 0
                    while True:
                        fmt = win32clipboard.EnumClipboardFormats(fmt)
                        if fmt == 0:
                            break
                        avail.append(fmt)
                    print(f"  Page {page_num}: No EMF. Available: {avail}")
            finally:
                win32clipboard.CloseClipboard()

        doc.Close(False)
    finally:
        word.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python word_emf_export.py <docx_path> <output_prefix>")
        sys.exit(1)
    export_emf(sys.argv[1], sys.argv[2])
