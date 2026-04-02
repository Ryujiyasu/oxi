"""Side-by-side GDI comparison: Oxi TextOutW vs Word EMF PlayEnhMetaFile.

Both rendered in the same DC for identical ClearType settings.

Usage: python side_by_side_gdi.py <docx_path> <emf_path> <output_png> [dpi]
"""
import subprocess
import ctypes
import ctypes.wintypes
import struct
import sys
import os

def main():
    if len(sys.argv) < 4:
        print("Usage: python side_by_side_gdi.py <docx_path> <emf_path> <output_png> [dpi]")
        sys.exit(1)

    docx_path = sys.argv[1]
    emf_path = sys.argv[2]
    output_png = sys.argv[3]
    dpi = int(sys.argv[4]) if len(sys.argv) > 4 else 150

    # Get Oxi layout
    oxi_root = os.path.join(os.path.dirname(__file__), '..', '..')
    result = subprocess.run(
        ['cargo', 'run', '--release', '--example', 'layout_json', '--', os.path.abspath(docx_path)],
        capture_output=True, text=True, errors='replace', cwd=oxi_root, timeout=120,
    )

    elements = []
    page_w, page_h = 612, 792
    for raw in result.stdout.split('\n'):
        line = raw.rstrip('\r')
        if line.startswith('PAGE\t'):
            parts = line.split('\t')
            page_w = float(parts[2])
            page_h = float(parts[3])
        elif line.startswith('TEXT\t'):
            parts = line.split('\t')
            elements.append({
                'type': 'text', 'x': float(parts[1]), 'y': float(parts[2]),
                'font_size': float(parts[5]), 'font': parts[6],
                'bold': parts[7] == '1', 'color': parts[11].strip(),
            })
        elif line.startswith('T\t') and elements and elements[-1]['type'] == 'text':
            elements[-1]['text'] = line[2:]
        elif line.startswith('BG\t'):
            parts = line.split('\t')
            elements.append({
                'type': 'bg', 'x': float(parts[1]), 'y': float(parts[2]),
                'width': float(parts[3]), 'height': float(parts[4]),
                'color': parts[5].strip(),
            })

    print(f"Oxi: {len(elements)} elements, page {page_w}x{page_h}pt")

    # Read EMF
    emf_data = open(os.path.abspath(emf_path), 'rb').read()
    # Parse EMF header for frame
    frame = struct.unpack_from('<iiii', emf_data, 24)
    frame_w_mm = (frame[2] - frame[0]) / 100.0
    frame_h_mm = (frame[3] - frame[1]) / 100.0
    frame_w_pt = frame_w_mm / 25.4 * 72.0
    frame_h_pt = frame_h_mm / 25.4 * 72.0
    print(f"EMF frame: {frame_w_pt:.1f}x{frame_h_pt:.1f}pt")

    # Render
    gdi32 = ctypes.windll.gdi32
    user32 = ctypes.windll.user32

    scale = dpi / 72.0
    pw = round(page_w * scale)
    ph = round(page_h * scale)
    # Two pages side by side
    total_w = pw * 2
    total_h = ph

    screen_dc = user32.GetDC(0)
    mem_dc = gdi32.CreateCompatibleDC(screen_dc)
    bitmap = gdi32.CreateCompatibleBitmap(screen_dc, total_w, total_h)
    old_bmp = gdi32.SelectObject(mem_dc, bitmap)

    class RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                     ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

    # White background
    white = gdi32.CreateSolidBrush(0x00FFFFFF)
    user32.FillRect(mem_dc, ctypes.byref(RECT(0, 0, total_w, total_h)), white)
    gdi32.DeleteObject(white)
    gdi32.SetBkMode(mem_dc, 1)  # TRANSPARENT

    # === LEFT: Word EMF ===
    gdi32.SetEnhMetaFileBits.restype = ctypes.c_void_p
    gdi32.SetEnhMetaFileBits.argtypes = [ctypes.c_uint, ctypes.c_char_p]
    hemf = gdi32.SetEnhMetaFileBits(len(emf_data), emf_data)
    # Map EMF content to margin area (gen_memo: left=90pt, top=72pt)
    margin_l = round(90.0 * scale)
    margin_t = round(72.0 * scale)
    content_w = round(frame_w_pt * scale)
    content_h = round(frame_h_pt * scale)
    play_rect = RECT(margin_l, margin_t, margin_l + content_w, margin_t + content_h)
    gdi32.PlayEnhMetaFile.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
    gdi32.PlayEnhMetaFile(mem_dc, hemf, ctypes.byref(play_rect))
    gdi32.DeleteEnhMetaFile.argtypes = [ctypes.c_void_p]
    gdi32.DeleteEnhMetaFile(hemf)

    # === RIGHT: Oxi TextOutW ===
    offset_x = pw  # right half
    for elem in elements:
        if elem['type'] == 'text' and 'text' in elem:
            c = elem['color'].lstrip('#')
            if len(c) == 6:
                r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
            else:
                r, g, b = 0, 0, 0
            gdi32.SetTextColor(mem_dc, r | (g << 8) | (b << 16))
            x = round(elem['x'] * scale) + offset_x
            y = round(elem['y'] * scale)
            fs = round(elem['font_size'] * scale)
            weight = 700 if elem['bold'] else 400
            family_buf = ctypes.create_unicode_buffer(elem['font'])
            font = gdi32.CreateFontW(-fs, 0, 0, 0, weight, 0, 0, 0, 1, 0, 0, 5, 0, family_buf)
            old_font = gdi32.SelectObject(mem_dc, font)
            text_buf = ctypes.create_unicode_buffer(elem['text'])
            gdi32.TextOutW(mem_dc, x, y, text_buf, len(elem['text']))
            gdi32.SelectObject(mem_dc, old_font)
            gdi32.DeleteObject(font)
        elif elem['type'] == 'bg':
            c = elem['color'].lstrip('#')
            if len(c) == 6:
                r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
            else:
                r, g, b = 0, 0, 0
            brush = gdi32.CreateSolidBrush(r | (g << 8) | (b << 16))
            x1 = round(elem['x'] * scale) + offset_x
            y1 = round(elem['y'] * scale)
            x2 = round((elem['x'] + elem['width']) * scale) + offset_x
            y2 = round((elem['y'] + elem['height']) * scale)
            user32.FillRect(mem_dc, ctypes.byref(RECT(x1, y1, x2, y2)), brush)
            gdi32.DeleteObject(brush)

    # Extract and save
    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [("biSize", ctypes.c_uint), ("biWidth", ctypes.c_long),
                     ("biHeight", ctypes.c_long), ("biPlanes", ctypes.c_ushort),
                     ("biBitCount", ctypes.c_ushort), ("biCompression", ctypes.c_uint),
                     ("biSizeImage", ctypes.c_uint), ("biXPelsPerMeter", ctypes.c_long),
                     ("biYPelsPerMeter", ctypes.c_long), ("biClrUsed", ctypes.c_uint),
                     ("biClrImportant", ctypes.c_uint)]
    class BITMAPINFO(ctypes.Structure):
        _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", ctypes.c_uint * 1)]

    bmi = BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = total_w
    bmi.bmiHeader.biHeight = -total_h
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32

    pixels = ctypes.create_string_buffer(total_w * total_h * 4)
    gdi32.GetDIBits(mem_dc, bitmap, 0, total_h, pixels, ctypes.byref(bmi), 0)

    from PIL import Image
    import numpy as np
    raw = np.frombuffer(pixels.raw, dtype=np.uint8).reshape(total_h, total_w, 4)
    rgb = raw[:, :, [2, 1, 0]]
    Image.fromarray(rgb).save(output_png)

    # Compare left vs right
    left = rgb[:, :pw, :]
    right = rgb[:, pw:, :]
    diff = np.abs(left.astype(int) - right.astype(int))
    n_diff = (diff.max(axis=2) > 0).sum()
    print(f"Saved {output_png} ({total_w}x{total_h})")
    print(f"Pixel diff (same DC): {n_diff}/{pw*ph} ({n_diff*100/(pw*ph):.3f}%)")

    # SSIM
    left_g = left.mean(axis=2)
    right_g = right.mean(axis=2)
    mu_a, mu_b = left_g.mean(), right_g.mean()
    var_a, var_b = left_g.var(), right_g.var()
    cov = ((left_g - mu_a) * (right_g - mu_b)).mean()
    c1, c2 = (0.01*255)**2, (0.03*255)**2
    ssim = (2*mu_a*mu_b+c1)*(2*cov+c2)/((mu_a**2+mu_b**2+c1)*(var_a+var_b+c2))
    print(f"SSIM (same DC): {ssim:.6f}")

    gdi32.SelectObject(mem_dc, old_bmp)
    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(0, screen_dc)

if __name__ == "__main__":
    main()
