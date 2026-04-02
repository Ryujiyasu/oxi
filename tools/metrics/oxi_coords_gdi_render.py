"""Render Oxi layout coordinates using GDI TextOutW (Python).

Takes Oxi layout_json output and renders with GDI TextOutW.
Produces a reference image using the exact same GDI pipeline as oxi-gdi-renderer,
but from Python — verifying that the coordinate calculation matches.

Usage: python oxi_coords_gdi_render.py <docx_path> <output_png> [dpi]
"""
import subprocess
import ctypes
import sys
import os

def get_oxi_layout(docx_path):
    """Run Oxi layout_json and parse TEXT/BG elements."""
    oxi_root = os.path.join(os.path.dirname(__file__), '..', '..')
    result = subprocess.run(
        ['cargo', 'run', '--release', '--example', 'layout_json', '--', docx_path],
        capture_output=True, text=True, errors='replace', cwd=oxi_root, timeout=120,
    )
    elements = []
    page_w, page_h = 612, 792
    for raw_line in result.stdout.split('\n'):
        line = raw_line.rstrip('\r')
        if line.startswith('PAGE\t'):
            parts = line.split('\t')
            page_w = float(parts[2])
            page_h = float(parts[3])
        elif line.startswith('TEXT\t'):
            parts = line.split('\t')
            x = float(parts[1])
            y = float(parts[2])
            width = float(parts[3])
            height = float(parts[4])
            font_size = float(parts[5])
            font_family = parts[6]
            bold = parts[7] == '1'
            color = parts[11].strip() if len(parts) > 11 else '#000000'
            elements.append({
                'type': 'text', 'x': x, 'y': y, 'width': width, 'height': height,
                'font_size': font_size, 'font': font_family, 'bold': bold, 'color': color,
            })
        elif line.startswith('T\t'):
            if elements and elements[-1]['type'] == 'text':
                elements[-1]['text'] = line[2:]
        elif line.startswith('BG\t'):
            parts = line.split('\t')
            x = float(parts[1])
            y = float(parts[2])
            width = float(parts[3])
            height = float(parts[4])
            color = parts[5].strip() if len(parts) > 5 else '#000000'
            elements.append({
                'type': 'border', 'x': x, 'y': y, 'width': width, 'height': height, 'color': color,
            })
    return elements, page_w, page_h


def parse_hex_color(color_str):
    c = color_str.strip().lstrip('#')
    if len(c) == 6:
        r = int(c[0:2], 16)
        g = int(c[2:4], 16)
        b = int(c[4:6], 16)
        return r, g, b
    return 0, 0, 0


def render_gdi(elements, page_w, page_h, dpi, output_png):
    gdi32 = ctypes.windll.gdi32
    user32 = ctypes.windll.user32

    TRANSPARENT = 1
    NONANTIALIASED_QUALITY = 3
    DEFAULT_CHARSET = 1

    scale = dpi / 72.0
    w = round(page_w * scale)
    h = round(page_h * scale)

    screen_dc = user32.GetDC(0)
    mem_dc = gdi32.CreateCompatibleDC(screen_dc)
    bitmap = gdi32.CreateCompatibleBitmap(screen_dc, w, h)
    old_bmp = gdi32.SelectObject(mem_dc, bitmap)

    class RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                     ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

    white_brush = gdi32.CreateSolidBrush(0x00FFFFFF)
    rect = RECT(0, 0, w, h)
    user32.FillRect(mem_dc, ctypes.byref(rect), white_brush)
    gdi32.DeleteObject(white_brush)
    gdi32.SetBkMode(mem_dc, TRANSPARENT)

    for elem in elements:
        if elem['type'] == 'border':
            r, g, b = parse_hex_color(elem['color'])
            rgb = r | (g << 8) | (b << 16)
            bw = max(1, round(elem['height'] * scale))
            x1 = round(elem['x'] * scale)
            y1 = round(elem['y'] * scale)
            x2 = round((elem['x'] + elem['width']) * scale)
            brush = gdi32.CreateSolidBrush(rgb)
            r2 = RECT(x1, y1, x2, y1 + bw)
            user32.FillRect(mem_dc, ctypes.byref(r2), brush)
            gdi32.DeleteObject(brush)
        elif elem['type'] == 'text' and 'text' in elem:
            r, g, b = parse_hex_color(elem['color'])
            rgb = r | (g << 8) | (b << 16)
            gdi32.SetTextColor(mem_dc, rgb)

            x = round(elem['x'] * scale)
            y = round(elem['y'] * scale)
            fs = round(elem['font_size'] * scale)
            weight = 700 if elem['bold'] else 400
            family = elem['font']
            family_buf = ctypes.create_unicode_buffer(family)
            font = gdi32.CreateFontW(
                -fs, 0, 0, 0, weight, 0, 0, 0,
                DEFAULT_CHARSET, 0, 0, NONANTIALIASED_QUALITY, 0, family_buf
            )
            old_font = gdi32.SelectObject(mem_dc, font)
            text = elem['text']
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
    bmi.bmiHeader.biHeight = -h
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32

    pixels = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mem_dc, bitmap, 0, h, pixels, ctypes.byref(bmi), 0)

    from PIL import Image
    import numpy as np
    raw = np.frombuffer(pixels.raw, dtype=np.uint8).reshape(h, w, 4)
    rgb_img = raw[:, :, [2, 1, 0]]
    Image.fromarray(rgb_img).save(output_png)
    print(f"Saved {output_png} ({w}x{h})")

    gdi32.SelectObject(mem_dc, old_bmp)
    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(0, screen_dc)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python oxi_coords_gdi_render.py <docx_path> <output_png> [dpi]")
        sys.exit(1)
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 150
    elements, pw, ph = get_oxi_layout(sys.argv[1])
    print(f"Got {len(elements)} elements, page {pw}x{ph}pt")
    render_gdi(elements, pw, ph, dpi, sys.argv[2])
