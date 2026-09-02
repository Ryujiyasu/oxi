"""Census: inline (`wp:inline`) shapes that CARRY TEXT (`w:txbxContent`).

These are the ones S839's inline path refuses (`tb.blocks.is_empty()`), so they
fall to the floating path and are placed by anchor instead of reserving height
in the line. Making them reserve height is the S1270 fix; this is its blast
radius.

Also counts the sub-population where the shape's paragraph carries a
`<w:br w:type="page"/>` -- the shape legal__02f84965dccfe4db is in, where the
box lands on the wrong side of the break.
"""
import re, sys, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

roots = [Path("pipeline_data/docx_corpus"), Path("tools/golden-test/documents/docx")]
n = with_inline_text = with_break = 0
rows = []
for root in roots:
    if not root.exists():
        continue
    for p in sorted(root.rglob("*.docx")):
        if p.name.startswith("~$"):
            continue
        n += 1
        try:
            xml = zipfile.ZipFile(p).read("word/document.xml").decode("utf-8", "replace")
        except Exception:
            continue
        hits = brk = 0
        for m in re.finditer(r"<wp:inline\b", xml):
            # the drawing's own extent ends at the matching </wp:inline>
            end = xml.find("</wp:inline>", m.start())
            if end < 0:
                continue
            seg = xml[m.start():end]
            if "<w:txbxContent" not in seg:
                continue
            hits += 1
            # does the HOST paragraph also carry an explicit page break?
            ps = max(xml.rfind("<w:p ", 0, m.start()), xml.rfind("<w:p>", 0, m.start()))
            pe = xml.find("</w:p>", end)
            if ps >= 0 and pe > 0 and '<w:br w:type="page"' in xml[end:pe]:
                brk += 1
        if hits:
            with_inline_text += 1
            rows.append((p.name, hits, brk))
        if brk:
            with_break += 1

print(f"{n} docx scanned")
print(f"  docs with an INLINE shape carrying text : {with_inline_text}")
print(f"  of those, the host paragraph also breaks: {with_break}")
print()
for name, hits, brk in sorted(rows, key=lambda r: -r[1])[:25]:
    print(f"  {name[:54]:<54} inline_text_shapes={hits:<3} with_page_break={brk}")
