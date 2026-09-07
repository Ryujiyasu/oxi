# -*- coding: utf-8 -*-
"""Dismiss Polaris Office's licence dialog (Cancel = keep the trial) and report
what windows are left, so the document window can be driven afterwards.

Usage: python tools/metrics/_polaris_dismiss.py [shot.png]
"""
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import win32con  # noqa: E402
import win32gui  # noqa: E402


def enum():
    out = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h):
            t = win32gui.GetWindowText(h)
            cls = win32gui.GetClassName(h)
            if t or cls.startswith("PO"):
                out.append((h, cls, t, win32gui.GetWindowRect(h)))
        return True

    win32gui.EnumWindows(cb, None)
    return out


def children(h):
    out = []

    def cb(c, _):
        out.append((c, win32gui.GetClassName(c), win32gui.GetWindowText(c), win32gui.GetWindowRect(c)))
        return True

    win32gui.EnumChildWindows(h, cb, None)
    return out


target = None
for h, cls, t, r in enum():
    if "POLARIS" in t.upper() or "Polaris" in cls or cls.startswith("PO"):
        print("window 0x%08X %-22s %-30s %s" % (h, cls, t[:30], r))
        for c, ccls, ct, cr in children(h):
            if ct:
                print("    child 0x%08X %-18s %-22s %s" % (c, ccls, ct[:22], cr))
        if "ライセンス" in t or "POLARIS OFFICE" in t.upper():
            target = h
if target:
    for c, ccls, ct, cr in children(target):
        if ct.strip() in ("キャンセル", "Cancel"):
            print("clicking", ct, "on 0x%08X" % c)
            win32gui.SendMessage(c, win32con.BM_CLICK, 0, 0)
            break
    time.sleep(3)
print("--- after ---")
for h, cls, t, r in enum():
    if cls.startswith("PO") or "Polaris" in t or ".docx" in t:
        print("window 0x%08X %-22s %-40s %s" % (h, cls, t[:40], r))
shot = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.environ["TMP"], "claude", "polaris_shot2.png")
from PIL import ImageGrab  # noqa: E402

ImageGrab.grab().save(shot)
print("shot ->", shot)
