"""Render Word document using GDI TextOutW with COM-extracted coordinates.

Extracts text positions from Word COM API, then renders using GDI TextOutW.
This produces the exact same GDI output as Word internally uses.

Usage: python word_gdi_render.py <docx_path> <output_png> [dpi]
"""
import win32com.client
import pythoncom
import ctypes
import struct
import sys
import os

def render_word_gdi(docx_path, output_png, dpi=150):
    docx_path = os.path.abspath(docx_path)
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True)

        # Get page size
        ps = doc.PageSetup
        page_w_pt = ps.PageWidth  # in points
        page_h_pt = ps.PageHeight

        scale = dpi / 72.0
        w_px = round(page_w_pt * scale)
        h_px = round(page_h_pt * scale)

        print(f"Page: {page_w_pt:.0f}x{page_h_pt:.0f}pt = {w_px}x{h_px}px at {dpi}DPI")

        # Extract text positions from Word COM
        texts = []
        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            rng = para.Range

            # Get each word/run position
            for wi in range(1, rng.Words.Count + 1):
                w_rng = rng.Words(wi)
                text = w_rng.Text.rstrip('\r\n\x07')
                if not text:
                    continue

                x = w_rng.Information(5)  # wdHorizontalPositionRelativeToPage
                y = w_rng.Information(6)  # wdVerticalPositionRelativeToPage

                fn = w_rng.Font.Name
                fs = w_rng.Font.Size
                bold = w_rng.Font.Bold
                # Resolve display color using Oxi's computed values
                # Word COM Font.Color is unreliable for theme colors (returns negative values).
                # Instead, use the Oxi layout output colors which correctly resolve themes.
                # For validation: use the Oxi layout color for Word GDI rendering too,
                # so color differences don't affect the pixel comparison.
                # This isolates position/font differences from color resolution differences.
                raw_color = w_rng.Font.Color
                if raw_color >= 0 and raw_color != -16777216:
                    r = raw_color & 0xFF
                    g = (raw_color >> 8) & 0xFF
                    b = (raw_color >> 16) & 0xFF
                else:
                    # Theme color: resolve via ObjectThemeColor + TintAndShade + ThemeColorScheme
                    try:
                        tc_obj = w_rng.Font.TextColor
                        otc = tc_obj.ObjectThemeColor
                        tint = tc_obj.TintAndShade
                        # Map msoThemeColor to ThemeColorScheme index
                        # msoThemeColorDark2=15 → scheme(3), msoThemeColorAccent1=4 → scheme(5)
                        otc_to_scheme = {1: 1, 2: 2, 3: 3, 4: 5, 5: 6, 6: 7, 7: 8, 8: 9, 9: 10,
                                        13: 1, 14: 2, 15: 3, 16: 4}
                        scheme_idx = otc_to_scheme.get(otc, 1)
                        base_rgb = doc.DocumentTheme.ThemeColorScheme(scheme_idx).RGB
                        br = base_rgb & 0xFF
                        bg = (base_rgb >> 8) & 0xFF
                        bb = (base_rgb >> 16) & 0xFF
                        if tint > 0:
                            r = min(255, int(br + (255 - br) * tint))
                            g = min(255, int(bg + (255 - bg) * tint))
                            b = min(255, int(bb + (255 - bb) * tint))
                        elif tint < 0:
                            r = max(0, int(br * (1 + tint)))
                            g = max(0, int(bg * (1 + tint)))
                            b = max(0, int(bb * (1 + tint)))
                        else:
                            r, g, b = br, bg, bb
                    except:
                        r, g, b = 0, 0, 0

                texts.append({
                    'text': text,
                    'x': round(x, 2),
                    'y': round(y, 2),
                    'font': fn,
                    'size': fs,
                    'bold': bold,
                    'r': r, 'g': g, 'b': b,
                })

        # Also extract borders (horizontal rules)
        # Check for bottom border on Title paragraph
        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            try:
                bb = para.Format.Borders(-3)  # wdBorderBottom
                if bb.LineStyle > 0:
                    # Get border position
                    y = para.Range.Information(6)
                    ls = para.Format.LineSpacing
                    lr = para.Format.LineSpacingRule
                    # Border is at bottom of paragraph
                    border_y = y + ls if lr == 0 else y + 20  # approximate
                    x_start = doc.PageSetup.LeftMargin
                    x_end = page_w_pt - doc.PageSetup.RightMargin
                    bw = bb.LineWidth / 8.0  # LineWidth in 1/8 pt
                    bc = bb.Color
                    if bc < 0:
                        br, bg, bb_c = 0, 0, 0
                    else:
                        br = bc & 0xFF
                        bg = (bc >> 8) & 0xFF
                        bb_c = (bc >> 16) & 0xFF
                    texts.append({
                        'type': 'border',
                        'x1': x_start, 'y1': border_y,
                        'x2': x_end, 'y2': border_y,
                        'width': bw,
                        'r': br, 'g': bg, 'b': bb_c,
                    })
            except:
                pass

        doc.Close(False)
        print(f"Extracted {len(texts)} text runs")

        # Now render using GDI
        render_gdi(texts, w_px, h_px, dpi, output_png)

    finally:
        word.Quit()


def render_gdi(texts, w, h, dpi, output_png):
    """Render extracted text positions using Windows GDI."""
    gdi32 = ctypes.windll.gdi32
    user32 = ctypes.windll.user32

    # GDI constants
    TRANSPARENT = 1
    CLEARTYPE_QUALITY = 5
    DEFAULT_CHARSET = 1
    SRCCOPY = 0x00CC0020

    screen_dc = user32.GetDC(0)
    mem_dc = gdi32.CreateCompatibleDC(screen_dc)
    bitmap = gdi32.CreateCompatibleBitmap(screen_dc, w, h)
    old_bmp = gdi32.SelectObject(mem_dc, bitmap)

    # White background
    class RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                     ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

    white_brush = gdi32.CreateSolidBrush(0x00FFFFFF)
    rect = RECT(0, 0, w, h)
    user32.FillRect(mem_dc, ctypes.byref(rect), white_brush)
    gdi32.DeleteObject(white_brush)

    gdi32.SetBkMode(mem_dc, TRANSPARENT)

    scale = dpi / 72.0

    for item in texts:
        if item.get('type') == 'border':
            # Draw horizontal line
            bw = max(1, round(item['width'] * scale))
            rgb = item['r'] | (item['g'] << 8) | (item['b'] << 16)
            pen = gdi32.CreatePen(0, bw, rgb)  # PS_SOLID=0
            old_pen = gdi32.SelectObject(mem_dc, pen)
            x1 = round(item['x1'] * scale)
            y1 = round(item['y1'] * scale)
            x2 = round(item['x2'] * scale)
            y2 = round(item['y2'] * scale)
            gdi32.MoveToEx(mem_dc, x1, y1, None)
            gdi32.LineTo(mem_dc, x2, y2)
            gdi32.SelectObject(mem_dc, old_pen)
            gdi32.DeleteObject(pen)
            continue

        x = round(item['x'] * scale)
        y = round(item['y'] * scale)
        fs = round(item['size'] * scale)
        weight = 700 if item['bold'] else 400
        rgb = item['r'] | (item['g'] << 8) | (item['b'] << 16)

        gdi32.SetTextColor(mem_dc, rgb)

        # Create font
        family = item['font']
        family_buf = ctypes.create_unicode_buffer(family)
        font = gdi32.CreateFontW(
            -fs, 0, 0, 0, weight,
            0, 0, 0,
            DEFAULT_CHARSET,
            0, 0,
            CLEARTYPE_QUALITY,
            0,
            family_buf
        )
        old_font = gdi32.SelectObject(mem_dc, font)

        # Draw text
        text = item['text']
        text_buf = ctypes.create_unicode_buffer(text)
        gdi32.TextOutW(mem_dc, x, y, text_buf, len(text))

        gdi32.SelectObject(mem_dc, old_font)
        gdi32.DeleteObject(font)

    # Extract pixels
    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", ctypes.c_uint), ("biWidth", ctypes.c_long),
            ("biHeight", ctypes.c_long), ("biPlanes", ctypes.c_ushort),
            ("biBitCount", ctypes.c_ushort), ("biCompression", ctypes.c_uint),
            ("biSizeImage", ctypes.c_uint), ("biXPelsPerMeter", ctypes.c_long),
            ("biYPelsPerMeter", ctypes.c_long), ("biClrUsed", ctypes.c_uint),
            ("biClrImportant", ctypes.c_uint),
        ]

    class BITMAPINFO(ctypes.Structure):
        _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", ctypes.c_uint * 1)]

    bmi = BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = w
    bmi.bmiHeader.biHeight = -h  # top-down
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32
    bmi.bmiHeader.biCompression = 0

    pixels = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mem_dc, bitmap, 0, h, pixels, ctypes.byref(bmi), 0)

    # Convert BGRA to RGB and save as PNG
    from PIL import Image
    import numpy as np
    raw = np.frombuffer(pixels.raw, dtype=np.uint8).reshape(h, w, 4)
    rgb = raw[:, :, [2, 1, 0]]  # BGR -> RGB
    img = Image.fromarray(rgb)
    img.save(output_png)
    print(f"Saved {output_png} ({w}x{h})")

    # Cleanup
    gdi32.SelectObject(mem_dc, old_bmp)
    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(0, screen_dc)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python word_gdi_render.py <docx_path> <output_png> [dpi]")
        sys.exit(1)
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 150
    render_word_gdi(sys.argv[1], sys.argv[2], dpi)
