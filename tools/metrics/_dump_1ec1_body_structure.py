# -*- coding: utf-8 -*-
"""Dump 1ec1's body structure to understand what's between body start and Shape 9."""
import sys, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc = z.read('word/document.xml').decode('utf-8')

body_m = re.search(r'<w:body>(.*?)</w:body>', doc, re.DOTALL)
body = body_m.group(1)
body_start_in_doc = body_m.start() + len('<w:body>')

# Find Shape 9 position in body
BOX5 = doc.find('□', 80000)
shape9_pos_in_body = BOX5 - body_start_in_doc
print(f"Body length: {len(body)}")
print(f"BOX5 in doc at pos {BOX5}; in body at offset {shape9_pos_in_body}")

# Walk body — find all top-level <w:p> and <w:tbl>, <w:sectPr>
# Top-level meaning direct children of <w:body>
def walk_top_level(content):
    """Return list of (tag, start, end) for top-level elements."""
    items = []
    pos = 0
    while pos < len(content):
        # Skip whitespace
        if content[pos] in '\n\r\t ':
            pos += 1
            continue
        # Look for next opening tag
        m = re.match(r'<(w:p|w:tbl|w:sectPr|mc:AlternateContent)([> ])', content[pos:])
        if m:
            tag = m.group(1)
            # Self-closing or paired?
            # Find matching close
            close_tag = f'</{tag}>'
            depth = 1
            search_from = pos + len(m.group(0))
            while depth > 0 and search_from < len(content):
                next_open = re.search(rf'<{re.escape(tag)}[ >]', content[search_from:])
                next_close = content.find(close_tag, search_from)
                if next_close < 0: break
                if next_open and next_open.start() + search_from < next_close:
                    depth += 1
                    search_from = next_open.start() + search_from + len(next_open.group(0))
                else:
                    depth -= 1
                    search_from = next_close + len(close_tag)
            items.append((tag, pos, search_from))
            pos = search_from
        else:
            pos += 1
    return items

top_items = walk_top_level(body)
print(f"\nTop-level body elements: {len(top_items)}")
for i, (tag, s, e) in enumerate(top_items):
    chunk = body[s:e]
    text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', chunk))[:60]
    contains_box = '□' in chunk
    contains_shape = '<wps:wsp' in chunk
    contains_alternate = '<mc:AlternateContent' in chunk
    marker = ''
    if s + len('<w:body>') >= shape9_pos_in_body:
        marker += ' <-- AT/AFTER Shape 9'
    if contains_box:
        marker += ' [HAS BOX]'
    if contains_shape:
        marker += ' [HAS SHAPE]'
    print(f"  [{i+1}] {tag}@{s}..{e} ({e-s}b) text={text!r}{marker}")
