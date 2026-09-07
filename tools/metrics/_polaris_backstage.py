# -*- coding: utf-8 -*-
"""Click a point on screen (the Polaris backstage is custom-drawn, so UI
Automation cannot see its items) and screenshot the result.

Usage: python tools/metrics/_polaris_backstage.py X Y [--wait 3] [--shot out.png]
"""
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from pywinauto import mouse  # noqa: E402

x, y = int(sys.argv[1]), int(sys.argv[2])
wait = 3.0
shot = os.path.join(os.environ["TMP"], "claude", "polaris_bs.png")
for i, a in enumerate(sys.argv):
    if a == "--wait":
        wait = float(sys.argv[i + 1])
    if a == "--shot":
        shot = sys.argv[i + 1]
print("click at", (x, y))
mouse.click(button="left", coords=(x, y))
time.sleep(wait)
from PIL import ImageGrab  # noqa: E402

ImageGrab.grab().save(shot)
print("shot ->", shot)
