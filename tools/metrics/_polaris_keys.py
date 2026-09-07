# -*- coding: utf-8 -*-
"""Send keystrokes to the focused Polaris Office document window and screenshot
the result.  Keys are given as a small vocabulary so a run is readable in the
log: ^p (ctrl+P), {ENTER}, {ESC}, {TAB}, {DOWN}, or literal text in quotes.

Usage: python tools/metrics/_polaris_keys.py "^p" [--wait 4] [--shot out.png]
"""
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import win32com.client  # noqa: E402
import win32gui  # noqa: E402

wait = 4.0
shot = os.path.join(os.environ["TMP"], "claude", "polaris_keys.png")
keys = []
i = 1
while i < len(sys.argv):
    a = sys.argv[i]
    if a == "--wait":
        i += 1
        wait = float(sys.argv[i])
    elif a == "--shot":
        i += 1
        shot = sys.argv[i]
    else:
        keys.append(a)
    i += 1

wins = []
win32gui.EnumWindows(lambda h, _: (wins.append((h, win32gui.GetWindowText(h))), True)[1], None)
target = next((h for h, t in wins if t.endswith(".docx") or t.endswith(".doc")), None)
dialogs = [(h, t) for h, t in wins if t and ("Polaris" in t or "印刷" in t or "保存" in t)]
print("doc window:", hex(target) if target else None, "| polaris dialogs:", [(hex(h), t) for h, t in dialogs])
if target:
    win32gui.SetForegroundWindow(target)
    time.sleep(0.8)
sh = win32com.client.Dispatch("WScript.Shell")
for k in keys:
    print("send %r" % k)
    sh.SendKeys(k)
    time.sleep(wait)
time.sleep(1.0)
from PIL import ImageGrab  # noqa: E402

ImageGrab.grab().save(shot)
print("shot ->", shot)
wins2 = []
win32gui.EnumWindows(lambda h, _: (wins2.append((h, win32gui.GetClassName(h), win32gui.GetWindowText(h))), True)[1], None)
for h, cls, t in wins2:
    if t and (t not in [x[1] for x in wins] or "印刷" in t or "保存" in t):
        print("new/notable window 0x%08X %-24s %r" % (h, cls, t[:50]))
