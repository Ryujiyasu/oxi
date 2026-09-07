# -*- coding: utf-8 -*-
"""Click a button by its caption in whatever Polaris Office dialog is up, then
screenshot.  Keeps the automation honest: nothing is clicked unless its caption
matches exactly one of the names given.

Usage: python tools/metrics/_polaris_click.py "いいえ(N)" [more captions...] [--shot out.png]
"""
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import win32con  # noqa: E402
import win32gui  # noqa: E402

args = [a for a in sys.argv[1:] if not a.startswith("--")]
shot = os.path.join(os.environ["TMP"], "claude", "polaris_click.png")
for i, a in enumerate(sys.argv):
    if a == "--shot" and i + 1 < len(sys.argv):
        shot = sys.argv[i + 1]
wanted = [a for a in args if a != shot]


def tops():
    out = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h):
            out.append((h, win32gui.GetClassName(h), win32gui.GetWindowText(h)))
        return True

    win32gui.EnumWindows(cb, None)
    return out


def kids(h):
    out = []

    def cb(c, _):
        out.append((c, win32gui.GetClassName(c), win32gui.GetWindowText(c)))
        return True

    try:
        win32gui.EnumChildWindows(h, cb, None)
    except Exception:
        pass
    return out


clicked = False
for h, cls, t in tops():
    if "Polaris" not in t and "POLARIS" not in t.upper():
        continue
    print("window 0x%08X %-22s %s" % (h, cls, t))
    for c, ccls, ct in kids(h):
        if ct:
            print("   child 0x%08X %-16s %r" % (c, ccls, ct))
        if ct.strip() in wanted:
            print("   -> clicking %r" % ct)
            win32gui.SetForegroundWindow(h)
            time.sleep(0.3)
            win32gui.SendMessage(c, win32con.BM_CLICK, 0, 0)
            clicked = True
            time.sleep(1.5)
            break
    if clicked:
        break
print("clicked:", clicked)
time.sleep(1.5)
from PIL import ImageGrab  # noqa: E402

ImageGrab.grab().save(shot)
print("shot ->", shot)
