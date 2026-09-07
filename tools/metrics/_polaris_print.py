# -*- coding: utf-8 -*-
"""Print the open Polaris Office document to PDF: focus the document window with
pywinauto (which handles the foreground restriction), send Ctrl+P, then report
what dialog came up so the next step can be aimed at real control names.

Usage: python tools/metrics/_polaris_print.py [--keys ^p] [--dump] [--shot out.png]
"""
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from pywinauto import Desktop  # noqa: E402

keys = "^p"
shot = os.path.join(os.environ["TMP"], "claude", "polaris_print.png")
dump = "--dump" in sys.argv
for i, a in enumerate(sys.argv):
    if a == "--keys":
        keys = sys.argv[i + 1]
    if a == "--shot":
        shot = sys.argv[i + 1]

desk = Desktop(backend="uia")
doc = None
for w in desk.windows():
    try:
        t = w.window_text()
    except Exception:
        continue
    if t.endswith(".docx") or t.endswith(".doc"):
        doc = w
        break
print("doc window:", doc.window_text() if doc else None)
if doc is None:
    raise SystemExit(1)
doc.set_focus()
time.sleep(1.0)
if keys:
    doc.type_keys(keys)
    time.sleep(5.0)
from PIL import ImageGrab  # noqa: E402

ImageGrab.grab().save(shot)
print("shot ->", shot)
for w in desk.windows():
    try:
        t = w.window_text()
    except Exception:
        continue
    if not t or t.endswith(".docx"):
        continue
    if any(k in t for k in ("印刷", "Print", "保存", "Polaris", "PDF")):
        print("dialog:", repr(t))
        if dump:
            def walk(c, d):
                if d > 4:
                    return
                for ch in c.children():
                    try:
                        n = ch.window_text()
                        ct = ch.element_info.control_type
                    except Exception:
                        continue
                    if n or ct in ("Button", "ComboBox", "Edit", "CheckBox"):
                        print("%s%-12s %r" % ("  " * d, ct, n[:44]))
                    walk(ch, d + 1)
            walk(w, 1)
