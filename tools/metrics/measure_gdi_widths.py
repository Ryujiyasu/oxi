"""Measure GDI ABC character widths for all ppem values.

Uses GetCharABCWidthsW to get pixel-accurate widths that match Word's GDI rendering.
Outputs JSON compatible with Oxi's gdi_full_widths.json format.

Usage: python measure_gdi_widths.py <font_name> [ppem_range] [output_json]
Example: python measure_gdi_widths.py Calibri 8-50
"""
import ctypes
import json
import sys

def measure_font(font_name, ppem_min=8, ppem_max=50):
    gdi32 = ctypes.windll.gdi32
    user32 = ctypes.windll.user32

    screen_dc = user32.GetDC(0)
    mem_dc = gdi32.CreateCompatibleDC(screen_dc)

    class ABC(ctypes.Structure):
        _fields_ = [('a', ctypes.c_int), ('b', ctypes.c_uint), ('c', ctypes.c_int)]

    result = {}
    for ppem in range(ppem_min, ppem_max + 1):
        family_buf = ctypes.create_unicode_buffer(font_name)
        font = gdi32.CreateFontW(-ppem, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, family_buf)
        old = gdi32.SelectObject(mem_dc, font)

        widths = {}
        for cp in range(32, 127):  # ASCII printable
            abc = ABC()
            gdi32.GetCharABCWidthsW(mem_dc, cp, cp, ctypes.byref(abc))
            total = abc.a + abc.b + abc.c
            widths[str(cp)] = total

        # Also common Unicode chars used in Japanese documents
        for cp in [0x3001, 0x3002, 0x3008, 0x3009, 0x300C, 0x300D, 0xFF08, 0xFF09,
                   0x2014, 0x2015, 0x2018, 0x2019, 0x201C, 0x201D, 0x2026]:
            abc = ABC()
            gdi32.GetCharABCWidthsW(mem_dc, cp, cp, ctypes.byref(abc))
            total = abc.a + abc.b + abc.c
            if total > 0:
                widths[str(cp)] = total

        result[str(ppem)] = widths
        gdi32.SelectObject(mem_dc, old)
        gdi32.DeleteObject(font)

    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(0, screen_dc)
    return result


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python measure_gdi_widths.py <font_name> [ppem_min-ppem_max]")
        sys.exit(1)

    font_name = sys.argv[1]
    ppem_range = sys.argv[2] if len(sys.argv) > 2 else "8-50"
    ppem_min, ppem_max = [int(x) for x in ppem_range.split('-')]

    print(f"Measuring {font_name} ppem {ppem_min}-{ppem_max}...")
    result = measure_font(font_name, ppem_min, ppem_max)

    # Merge with existing gdi_full_widths.json if exists
    output_path = f"pipeline_data/gdi_widths_{font_name}.json"
    with open(output_path, 'w') as f:
        json.dump(result, f)

    total_chars = sum(len(v) for v in result.values())
    print(f"Saved {output_path}: {len(result)} ppems, {total_chars} total width entries")
