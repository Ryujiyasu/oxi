# -*- coding: utf-8 -*-
"""Drive Polaris Office through UI Automation instead of keystrokes, so nothing
depends on which window happens to hold the foreground.

Sub-commands:
  dump [depth]        list the ribbon/dialog controls under the document window
  click <name>        invoke the control whose name matches (exact, then prefix)
  save-pdf <docx> <out.pdf>   open, export to PDF, close

Usage: python tools/metrics/_polaris_uia.py dump 3
"""
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from pywinauto import Application, Desktop  # noqa: E402

EXE = r"C:\Program Files (x86)\Polaris Office\Office11\Binary\PWord_PC11.exe"


def doc_window():
    for w in Desktop(backend="uia").windows():
        try:
            t = w.window_text()
        except Exception:
            continue
        if t.endswith(".docx") or t.endswith(".doc"):
            return w
    return None


def walk(ctrl, depth, max_depth, seen=0):
    if depth > max_depth:
        return seen
    for c in ctrl.children():
        try:
            name = c.window_text()
            ct = c.element_info.control_type
        except Exception:
            continue
        if name or ct in ("Button", "TabItem", "MenuItem"):
            print("%s%-14s %r" % ("  " * depth, ct, name[:48]))
            seen += 1
        seen = walk(c, depth + 1, max_depth, seen)
    return seen


cmd = sys.argv[1] if len(sys.argv) > 1 else "dump"
if cmd == "dump":
    w = doc_window()
    print("doc window:", w.window_text() if w else None)
    if w:
        walk(w, 0, int(sys.argv[2]) if len(sys.argv) > 2 else 3)
elif cmd == "click":
    want = sys.argv[2]
    w = doc_window()
    hit = None
    def find(ctrl, depth):
        global hit
        if hit or depth > 6:
            return
        for c in ctrl.children():
            try:
                name = c.window_text()
            except Exception:
                continue
            if name == want:
                hit = c
                return
            find(c, depth + 1)
    find(w, 0)
    print("found:", hit.element_info.control_type if hit else None)
    if hit:
        try:
            hit.invoke()
        except Exception:
            hit.click_input()
        time.sleep(2)
    from PIL import ImageGrab
    shot = os.path.join(os.environ["TMP"], "claude", "polaris_uia.png")
    ImageGrab.grab().save(shot)
    print("shot ->", shot)
