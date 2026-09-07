# -*- coding: utf-8 -*-
"""Open one .docx in Polaris Office's word module and screenshot it, so the
window layout can be read before any keystroke is scripted.

Usage: python tools/metrics/_polaris_open.py <docx> [seconds] [shot.png]
"""
import os
import subprocess
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
EXE = r"C:\Program Files (x86)\Polaris Office\Office11\Binary\PWord_PC11.exe"
docx = os.path.abspath(sys.argv[1])
wait = float(sys.argv[2]) if len(sys.argv) > 2 else 25.0
shot = sys.argv[3] if len(sys.argv) > 3 else os.path.join(os.environ["TMP"], "claude", "polaris_shot.png")
p = subprocess.Popen([EXE, docx])
print("launched pid", p.pid, "->", os.path.basename(docx))
time.sleep(wait)
import win32gui  # noqa: E402


def windows():
    out = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h):
            t = win32gui.GetWindowText(h)
            if t:
                out.append((h, t, win32gui.GetWindowRect(h)))
        return True

    win32gui.EnumWindows(cb, None)
    return out


for h, t, r in windows():
    if any(k in t for k in ("Polaris", os.path.basename(docx)[:12], "PDF", "印刷")) or r[2] - r[0] > 500:
        print("win 0x%08X %-40s %s" % (h, t[:40], r))
from PIL import ImageGrab  # noqa: E402

os.makedirs(os.path.dirname(shot), exist_ok=True)
ImageGrab.grab().save(shot)
print("shot ->", shot)
