"""Capture Word Desktop rendering as screenshot.

Opens document in Word, sets zoom to exact DPI match, captures window content.
This gives the EXACT pixels Word Desktop displays to the user.

Usage: python word_screenshot.py <docx_path> <output_png> [dpi]
"""
import win32com.client
import win32gui
import win32ui
import win32con
import pythoncom
import ctypes
import time
import sys
import os

def capture_word(docx_path, output_png, target_dpi=150):
    docx_path = os.path.abspath(docx_path)
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = True
    word.WindowState = 1  # wdWindowStateNormal

    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        win = word.ActiveWindow
        view = win.View
        view.Type = 3  # wdPrintView

        # Set zoom to match target DPI
        # Word zoom 100% = 96 DPI (Windows default)
        # For 150 DPI output: zoom = 150/96 * 100 = 156.25%
        zoom_pct = round(target_dpi / 96.0 * 100)
        win.View.Zoom.Percentage = zoom_pct
        print(f"Set zoom to {zoom_pct}%")

        # Wait for rendering
        time.sleep(2)

        # Get Word window handle
        hwnd = win.Hwnd
        print(f"Word HWND: {hwnd}")

        # Get window client rect
        rect = win32gui.GetClientRect(hwnd)
        w = rect[2] - rect[0]
        h = rect[3] - rect[1]
        print(f"Client rect: {w}x{h}")

        # Capture using PrintWindow (works even if occluded)
        hwnd_dc = win32gui.GetDC(hwnd)
        mem_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mem_dc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mem_dc, w, h)
        save_dc.SelectObject(bmp)

        # PrintWindow captures the actual rendered content
        ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 3)  # PW_RENDERFULLCONTENT

        # Save bitmap
        bmp.SaveBitmapFile(save_dc, output_png.replace('.png', '.bmp'))

        # Convert to PNG
        from PIL import Image
        img = Image.open(output_png.replace('.png', '.bmp'))
        img.save(output_png)
        os.remove(output_png.replace('.png', '.bmp'))
        print(f"Saved {output_png} ({img.size[0]}x{img.size[1]})")

        save_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)

        doc.Close(False)
    finally:
        word.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python word_screenshot.py <docx_path> <output_png> [dpi]")
        sys.exit(1)
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 150
    capture_word(sys.argv[1], sys.argv[2], dpi)
