# -*- coding: utf-8 -*-
"""Measure per-paragraph Y (Information(6), collapsed start range = R30 fix)
for the mult_grid repro, and report each test paragraph's advance."""
import os, sys, time, shutil, tempfile
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.join(REPO, "tools", "golden-test", "repros", "mult_grid", "mult_grid.docx")

wd = win32com.client.gencache.EnsureDispatch("Word.Application")
wd.Visible = False
fd, tmp = tempfile.mkstemp(suffix=".docx", prefix="mg_"); os.close(fd)
shutil.copy(DOCX, tmp)
doc = wd.Documents.Open(tmp, ReadOnly=True)
time.sleep(0.3)
rows = []
try:
    n = doc.Paragraphs.Count
    for pi in range(1, n + 1):
        p = doc.Paragraphs(pi)
        rng = p.Range
        txt = (rng.Text or "").rstrip("\r\a\x07")
        cs = doc.Range(rng.Start, rng.Start)
        y = cs.Information(6)   # wdVerticalPositionRelativeToPage (points)
        pg = cs.Information(3)  # wdActiveEndPageNumber
        rows.append((pi, pg, round(y, 2), txt[:16]))
finally:
    doc.Close(False)
    wd.Quit()
    try: os.remove(tmp)
    except: pass

print(f"{'i':>3} {'pg':>2} {'y':>8}  text")
for pi, pg, y, t in rows:
    print(f"{pi:>3} {pg:>2} {y:>8}  {t}")

# advances: a test para sits between anchorA (i-1) and anchorB (i+1)
print("\n--- test paragraph advances (Y(next) - Y(this)) ---")
for idx in range(1, len(rows) - 1):
    pi, pg, y, t = rows[idx]
    if t.startswith("ANCHOR"):
        continue
    nxt = rows[idx + 1]
    if nxt[1] != pg:
        print(f"  {t:16} advance=PAGE-BREAK (this p{pg} next p{nxt[1]})")
    else:
        print(f"  {t:16} advance={nxt[2]-y:7.2f}  (this_y={y} next_y={nxt[2]})")
