"""Word GDI Direct Renderer: get text positions from Word COM, render with GDI.
Eliminates EMF coordinate mapping issues by using COM positions directly.
Both this and oxi-gdi-renderer use TextOutW → font rendering is identical,
only layout positions differ.

Usage: python word_gdi_direct.py <input.docx> <output_prefix> [dpi]
"""
import win32com.client
import ctypes
from ctypes import wintypes
import sys
import os
import time

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

def render_word_gdi(docx_path, output_prefix, dpi=150):
    # Open document in Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    time.sleep(1)

    page_count = doc.ComputeStatistics(2)
    ps = doc.PageSetup
    page_w_pt = ps.PageWidth
    page_h_pt = ps.PageHeight
    print(f"Pages: {page_count}, size: {page_w_pt:.1f}x{page_h_pt:.1f}pt, DPI: {dpi}")

    scale = dpi / 72.0
    w = int(round(page_w_pt * scale))
    h = int(round(page_h_pt * scale))

    # Setup GDI
    screen_dc = user32.GetDC(0)

    for page_num in range(1, page_count + 1):
        mem_dc = gdi32.CreateCompatibleDC(screen_dc)
        bitmap = gdi32.CreateCompatibleBitmap(screen_dc, w, h)
        old_bmp = gdi32.SelectObject(mem_dc, bitmap)

        # White background
        white_brush = gdi32.CreateSolidBrush(0x00FFFFFF)
        rect = (ctypes.c_long * 4)(0, 0, w, h)
        user32.FillRect(mem_dc, rect, white_brush)
        gdi32.DeleteObject(white_brush)
        gdi32.SetBkMode(mem_dc, 1)  # TRANSPARENT

        # Collect text from all paragraphs on this page
        chars_rendered = 0
        for pi in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(pi)
            rng = p.Range

            # Check if this paragraph is on the current page
            try:
                page = rng.Information(1)  # wdActiveEndPageNumber
            except:
                continue
            if page != page_num:
                if page > page_num:
                    break
                continue

            # Get characters with positions
            chars = rng.Characters
            n = chars.Count

            for ci in range(1, n + 1):
                try:
                    c = chars(ci)
                    ch = c.Text
                    if ch in ('\r', '\n', '\x07', '\x0C', '\x0B'):
                        continue

                    x_pt = c.Information(5)  # wdHorizontalPositionRelativeToPage
                    y_pt = c.Information(6)  # wdVerticalPositionRelativeToPage

                    font_name = c.Font.Name
                    font_size = c.Font.Size
                    bold = c.Font.Bold
                    color_val = c.Font.Color  # RGB as long

                    if font_size <= 0 or font_size > 200:
                        continue

                    # Convert to pixels
                    x_px = int(round(x_pt * scale))
                    y_px = int(round(y_pt * scale))
                    fs_px = int(round(font_size * scale))

                    # Parse color
                    if color_val == -16777216 or color_val == 0:  # wdColorAutomatic or black
                        rgb = 0x00000000
                    else:
                        r = color_val & 0xFF
                        g = (color_val >> 8) & 0xFF
                        b = (color_val >> 16) & 0xFF
                        rgb = r | (g << 8) | (b << 16)

                    gdi32.SetTextColor(mem_dc, rgb)

                    # Create font
                    weight = 700 if bold else 400
                    font_buf = ctypes.create_unicode_buffer(font_name)
                    hfont = gdi32.CreateFontW(
                        -fs_px, 0, 0, 0, weight,
                        0, 0, 0, 1, 0, 0, 5, 0, font_buf
                    )
                    old_font = gdi32.SelectObject(mem_dc, hfont)

                    # Draw character
                    text_buf = ctypes.create_unicode_buffer(ch)
                    gdi32.TextOutW(mem_dc, x_px, y_px, text_buf, len(ch))

                    gdi32.SelectObject(mem_dc, old_font)
                    gdi32.DeleteObject(hfont)
                    chars_rendered += 1

                except Exception as e:
                    continue

        print(f"  Page {page_num}: {chars_rendered} chars rendered")

        # Extract bitmap pixels
        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [
                ('biSize', ctypes.c_uint32), ('biWidth', ctypes.c_int32),
                ('biHeight', ctypes.c_int32), ('biPlanes', ctypes.c_uint16),
                ('biBitCount', ctypes.c_uint16), ('biCompression', ctypes.c_uint32),
                ('biSizeImage', ctypes.c_uint32), ('biXPelsPerMeter', ctypes.c_int32),
                ('biYPelsPerMeter', ctypes.c_int32), ('biClrUsed', ctypes.c_uint32),
                ('biClrImportant', ctypes.c_uint32),
            ]

        bmi = BITMAPINFOHEADER()
        bmi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.biWidth = w
        bmi.biHeight = -h  # top-down
        bmi.biPlanes = 1
        bmi.biBitCount = 32

        pixels = (ctypes.c_uint8 * (w * h * 4))()
        gdi32.GetDIBits(mem_dc, bitmap, 0, h, pixels, ctypes.byref(bmi), 0)

        # Convert BGRA to RGB and save as PNG
        import numpy as np
        from PIL import Image
        arr = np.frombuffer(pixels, dtype=np.uint8).reshape(h, w, 4)
        rgb = arr[:, :, [2, 1, 0]]  # BGRA → RGB

        out_path = f"{output_prefix}_p{page_num}.png"
        Image.fromarray(rgb).save(out_path)
        print(f"  Saved {out_path} ({w}x{h})")

        gdi32.SelectObject(mem_dc, old_bmp)
        gdi32.DeleteObject(bitmap)
        gdi32.DeleteDC(mem_dc)

    user32.ReleaseDC(0, screen_dc)
    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python word_gdi_direct.py <input.docx> <output_prefix> [dpi]")
        sys.exit(1)
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 150
    render_word_gdi(sys.argv[1], sys.argv[2], dpi)
