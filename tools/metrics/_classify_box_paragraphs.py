# -*- coding: utf-8 -*-
"""Definitively classify each □ paragraph in 1ec1: in-shape vs body."""
import sys, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc = z.read('word/document.xml').decode('utf-8')

BOX = '□'
positions = [m.start() for m in re.finditer(BOX, doc)]
print(f"Total □ in document.xml: {len(positions)}")

# For each position, find the immediate enclosing AlternateContent (shape) or body.
# The correct check: look back to <mc:AlternateContent and forward to </mc:AlternateContent.
# If a containing AC pair is found, the □ is inside a shape.

def in_shape(pos):
    """Find smallest <mc:AlternateContent>...</mc:AlternateContent> pair containing pos."""
    # Find all AC starts before pos and ends after pos
    starts = [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc) if m.start() < pos]
    ends = [m.end() for m in re.finditer(r'</mc:AlternateContent>', doc) if m.end() > pos]
    if not starts or not ends: return None
    # Smallest enclosing pair
    enclosing = None
    for s in reversed(starts):
        for e in ends:
            if e > s and e > pos:
                enclosing = (s, e)
                break
        if enclosing: break
    if not enclosing: return None
    s, e = enclosing
    # Verify pos is between
    if s < pos < e:
        # Get docPr id
        block = doc[s:e]
        m = re.search(r'<wp:docPr[^>]*?id="(\d+)"[^>]*?name="([^"]*)"', block)
        if m: return (m.group(1), m.group(2))
    return None

print(f"\n{'idx':>3} {'pos':>7} {'context':<30} {'text(50)':<60}")
for i, pos in enumerate(positions):
    shape = in_shape(pos)
    # Get paragraph text
    p_start = max(doc.rfind('<w:p ', 0, pos), doc.rfind('<w:p>', 0, pos))
    p_end = doc.find('</w:p>', pos) + len('</w:p>')
    para = doc[p_start:p_end]
    text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para))[:50]
    # Get ind from this paragraph
    ind_left = re.search(r'<w:ind[^>]*?w:left="(\d+)"', para)
    ind_first = re.search(r'<w:ind[^>]*?w:firstLine="(-?\d+)"', para)
    ind_str = f'left={ind_left.group(1)}' if ind_left else ''
    if ind_first: ind_str += f' firstLine={ind_first.group(1)}'
    if not ind_str: ind_str = 'no_ind'
    ctx = f"Shape id={shape[0]}" if shape else "BODY"
    print(f"  {i+1:>3} {pos:>7} {ctx:<30} {text!r}  [{ind_str}]")
