"""Scan all 49 docs for w:characterSpacingControl setting."""
import zipfile
import re
from pathlib import Path

values = {}
import sys
target_dir = sys.argv[1] if len(sys.argv) > 1 else "pipeline_data/docx"
for d in sorted(Path(target_dir).glob("*.docx")):
    if d.name.startswith("~$"):
        continue
    try:
        with zipfile.ZipFile(d) as z:
            try:
                s = z.read("word/settings.xml").decode("utf-8")
            except KeyError:
                values.setdefault("(no settings.xml)", []).append(d.stem)
                continue
        m = re.search(r'<w:characterSpacingControl w:val="([^"]+)"', s)
        if m:
            values.setdefault(m.group(1), []).append(d.stem)
        else:
            values.setdefault("(absent)", []).append(d.stem)
    except Exception as e:
        print(f"err {d.name}: {e}")

print(f"Total docs: {sum(len(v) for v in values.values())}\n")
for v, docs in sorted(values.items()):
    print(f"{v}: {len(docs)}")
    for d in docs[:5]:
        print(f"  {d}")
    if len(docs) > 5:
        print(f"  ... and {len(docs)-5} more")
