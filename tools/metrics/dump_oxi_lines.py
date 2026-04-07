"""Dump Oxi line content (text per line) for a docx using full layout output (TSV)."""
import subprocess
import sys

docx = sys.argv[1]
result = subprocess.run(
    ["cargo", "run", "--release", "--example", "layout_json", "--", docx],
    capture_output=True, text=True, errors="replace", timeout=120,
)

# Parse TSV: lines starting with TEXT or T\t
# Format: TEXT\tx\ty\tfontSize\twidth\theight\tfont\t...
# Then: T\t<glyph>
current = None  # (x, y, accumulated_text)
lines = {}  # (page, para, y) -> list of (x, text)
page = 0
para_idx = 0  # rough counter; we don't have explicit para info in this format
# Use Y change as paragraph proxy
prev_y = None

for raw in result.stdout.splitlines():
    parts = raw.split("\t")
    if parts[0] == "PAGE":
        page = int(parts[1])
    elif parts[0] == "TEXT":
        # Flush prev
        if current is not None:
            x, y, text = current
            lines.setdefault((page, round(y * 2) / 2), []).append((x, text))
        x = float(parts[1])
        y = float(parts[2])
        current = [x, y, ""]
    elif parts[0] == "T" and current is not None:
        current[2] += parts[1]

if current is not None:
    x, y, text = current
    lines.setdefault((page, round(y * 2) / 2), []).append((x, text))

# Group consecutive Ys into paragraph-ish groups by Y gap
prev_key = None
for key in sorted(lines.keys()):
    pg, y = key
    frags = sorted(lines[key])
    text = "".join(t for _, t in frags)
    sys.stdout.buffer.write(
        (f"PG{pg} y={y}: ({len(text)}ch) " + text + "\n").encode("utf-8", errors="replace")
    )
