# -*- coding: utf-8 -*-
"""Export one .docx to PDF through Polaris Office's own engine, unattended.

The backstage pane is custom-drawn (UI Automation cannot see its items), so the
export is driven by the ribbon's Application menu plus fixed points inside that
pane; every point is verified by a screenshot when --shots is given.  The window
is maximised to 1920x1152 by Polaris itself, which is what the coordinates below
assume.

Usage: python tools/metrics/_polaris_export.py <docx> <out.pdf> [--shots]
"""
import os
import subprocess
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import win32con  # noqa: E402
import win32gui  # noqa: E402
from pywinauto import Desktop, keyboard, mouse  # noqa: E402
from PIL import ImageGrab  # noqa: E402

EXE = r"C:\Program Files (x86)\Polaris Office\Office11\Binary\PWord_PC11.exe"
P_EXPORT_TAB = (50, 288)     # backstage: エクスポート
P_PDF_ITEM = (343, 120)      # export pane: Adobe PDF ファイル (.pdf)
P_EXPORT_BTN = (886, 368)    # export pane: エクスポート button
P_FILENAME = (540, 457)      # save dialog: ファイル名 edit
P_SAVE = (790, 561)          # save dialog: 保存(S)
SHOTS = "--shots" in sys.argv
docx = os.path.abspath(sys.argv[1])
out = os.path.abspath(sys.argv[2]).replace("/", "\\")
os.makedirs(os.path.dirname(out), exist_ok=True)
if os.path.exists(out):
    os.remove(out)


def shot(tag):
    if SHOTS:
        p = os.path.join(os.environ["TMP"], "claude", "polaris_%s.png" % tag)
        ImageGrab.grab().save(p)
        print("   shot", p)


def click_button(caption):
    """Click a button by caption in any visible Polaris dialog."""
    found = [False]

    def cb(h, _):
        if not win32gui.IsWindowVisible(h):
            return True
        t = win32gui.GetWindowText(h)
        if "Polaris" not in t and "POLARIS" not in t.upper():
            return True

        def kid(c, _):
            if win32gui.GetWindowText(c).strip() == caption:
                win32gui.SendMessage(c, win32con.BM_CLICK, 0, 0)
                found[0] = True
            return True

        win32gui.EnumChildWindows(h, kid, None)
        return True

    win32gui.EnumWindows(cb, None)
    return found[0]


def doc_window():
    for w in Desktop(backend="uia").windows():
        try:
            t = w.window_text()
        except Exception:
            continue
        if t.endswith(".docx") or t.endswith(".doc"):
            return w
    return None


proc = subprocess.Popen([EXE, docx])
print("launched", os.path.basename(docx), "pid", proc.pid)
time.sleep(28)
if click_button("いいえ(&N)"):
    print("   declined the default-program prompt")
    time.sleep(2)
w = doc_window()
if w is None:
    print("no document window"); raise SystemExit(2)
w.set_focus()
time.sleep(1.0)

appmenu = [None]


def find(c, d):
    if appmenu[0] or d > 4:
        return
    for ch in c.children():
        try:
            n = ch.window_text()
        except Exception:
            continue
        if n == "Application menu":
            appmenu[0] = ch
            return
        find(ch, d + 1)


find(w, 0)
if appmenu[0] is None:
    print("no application menu"); raise SystemExit(3)
appmenu[0].invoke()
time.sleep(2.5)
shot("bs")
mouse.click(button="left", coords=P_EXPORT_TAB)
time.sleep(3.0)
mouse.click(button="left", coords=P_PDF_ITEM)
time.sleep(3.0)
shot("pdfpane")
mouse.click(button="left", coords=P_EXPORT_BTN)
time.sleep(5.0)
mouse.click(button="left", coords=P_FILENAME)
time.sleep(0.6)
keyboard.send_keys("^a")
time.sleep(0.3)
keyboard.send_keys(out.replace("(", "{(}").replace(")", "{)}"), with_spaces=True, pause=0.01)
time.sleep(0.6)
shot("savedlg")
mouse.click(button="left", coords=P_SAVE)
for _ in range(40):
    time.sleep(2.0)
    if os.path.exists(out) and os.path.getsize(out) > 1000:
        break
    click_button("はい(&Y)")  # overwrite prompt, if any
print("exported:", os.path.exists(out), os.path.getsize(out) if os.path.exists(out) else "")
shot("done")
subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], capture_output=True)
time.sleep(2)
