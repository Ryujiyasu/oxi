"""Compare Word and Oxi char positions for the test yakumono doc."""
import win32com.client
import os
import sys
import time
import subprocess

docx = os.path.abspath("pipeline_data/docx_test/test_yakumono_compressed.docx")

# Word side
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(docx, ReadOnly=True)
time.sleep(1)

p = doc.Paragraphs(1)
chars = p.Range.Characters
print("=== Word ===")
for ci in range(1, chars.Count + 1):
    try:
        c = chars(ci)
        ch = c.Text
        if ch in ("\r", "\x07"):
            continue
        cx = c.Information(5)
        sys.stdout.buffer.write(f"  '{ch}' x={cx:.2f}\n".encode("utf-8", "replace"))
    except:
        continue
doc.Close(SaveChanges=False)
word.Quit()

# Oxi side
print("\n=== Oxi ===")
result = subprocess.run(
    ["cargo", "run", "--release", "--example", "layout_json", "--", docx],
    capture_output=True, text=True, errors="replace", timeout=120,
)
y_first = None
for raw in result.stdout.splitlines():
    p = raw.split("\t")
    if p[0] == "TEXT":
        y = float(p[2])
        if y_first is None:
            y_first = y
        if y == y_first:
            cur_x = float(p[1])
    elif p[0] == "T" and y_first is not None:
        sys.stdout.buffer.write(f"  '{p[1]}' x={cur_x:.2f}\n".encode("utf-8", "replace"))
