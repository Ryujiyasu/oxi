# -*- coding: utf-8 -*-
import sys, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc = z.read('word/document.xml').decode('utf-8')

BOX = '□'
positions = [m.start() for m in re.finditer(BOX, doc)]
pos1 = positions[0]
pos5 = positions[4]  # □３

# Walk back to enclosing structure: look for <w:tbl, <w:sdt, <w:tc, w:p, w:txbxContent
def parents_in(haystack, end_pos, opens):
    """Find all opening tags in opens that are still open at end_pos (i.e., open before end_pos and close after end_pos)."""
    open_pos = []
    for tag in opens:
        # Find all opens before end_pos
        opens_at = [m.start() for m in re.finditer(rf'<{re.escape(tag)}[ >]', haystack[:end_pos])]
        closes_at = [m.start() for m in re.finditer(rf'</{re.escape(tag)}>', haystack[:end_pos])]
        if len(opens_at) > len(closes_at):
            # Some still open
            stack_count = len(opens_at) - len(closes_at)
            open_pos.append((tag, stack_count, opens_at[-stack_count]))
    return open_pos

print(f"=== Context around BOX[1] pos={pos1} ===")
parents = parents_in(doc, pos1, ['w:tbl', 'w:tc', 'w:tr', 'w:sdt', 'w:txbxContent', 'wps:wsp', 'mc:AlternateContent'])
for t, n, p in parents:
    print(f"  Open: <{t}> (depth={n}) opened at pos {p}")

# Show 500 chars BEFORE BOX[1]
print("\n--- 500 chars before BOX[1] ---")
print(doc[max(0, pos1-500):pos1])
print("\n--- 200 chars after BOX[1] ---")
print(doc[pos1:pos1+200])

print("\n\n=== Context around BOX[5] (□３) pos={} ===".format(pos5))
parents = parents_in(doc, pos5, ['w:tbl', 'w:tc', 'w:tr', 'w:sdt', 'w:txbxContent', 'wps:wsp', 'mc:AlternateContent'])
for t, n, p in parents:
    print(f"  Open: <{t}> (depth={n}) opened at pos {p}")
