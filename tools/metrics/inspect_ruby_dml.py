"""Inspect Word DML cache and Oxi structure for ruby docs side-by-side."""
import json
import sys
import subprocess
import os

doc = sys.argv[1]  # ruby_text_lineheight_11
docx = f"pipeline_data/docx/{doc}.docx"
cache = f"pipeline_data/word_dml/{doc}.json"

# Word side
with open(cache, encoding="utf-8") as f:
    d = json.load(f)
print("=== Word DML ===")
for p in d["paragraphs"][:3]:
    print(f"P{p['index']} y={p['y']}")
    text = p.get("text", "")
    sys.stdout.buffer.write(("  text=" + repr(text) + "\n").encode("utf-8", errors="replace"))
    for li, ln in enumerate(p.get("lines", [])):
        ltext = ln.get("text", "")
        sys.stdout.buffer.write(
            (f"  L{li}: y={ln['y']} chars={ln.get('chars','?')} text=" + repr(ltext) + "\n").encode("utf-8", errors="replace")
        )

# Oxi side
print("\n=== Oxi structure ===")
result = subprocess.run(
    ["cargo", "run", "--release", "--example", "layout_json", "--", docx, "--structure"],
    capture_output=True, text=True, errors="replace", timeout=120,
)
for line in result.stdout.splitlines():
    if line.startswith("PAGE") or line.startswith("PARA") or line.startswith("  LINE"):
        print(line)
