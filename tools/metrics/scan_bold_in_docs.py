"""Definitive scan: count actual bold runs (run-level rPr) per doc."""
import zipfile
import re
from pathlib import Path

p = re.compile(r"<w:r\b[^>]*>.*?</w:r>", re.DOTALL)
b_re = re.compile(r"<w:b\s*/>|<w:b\s+w:val=\"1\"")

docs = sorted(Path("pipeline_data/docx").glob("*.docx"))
hits = []
for d in docs:
    try:
        with zipfile.ZipFile(d) as z:
            x = z.read("word/document.xml").decode("utf-8")
        # Look for bold inside w:rPr
        run_count = 0
        for m in re.finditer(r"<w:r\b[^>]*>.*?</w:r>", x, re.DOTALL):
            run_xml = m.group(0)
            # Find rPr
            rpr_m = re.search(r"<w:rPr>(.*?)</w:rPr>", run_xml, re.DOTALL)
            if rpr_m and b_re.search(rpr_m.group(1)):
                run_count += 1
        if run_count > 0:
            hits.append((d.name, run_count))
    except Exception as e:
        print(f"err {d.name}: {e}")

print(f"Total docs with bold runs: {len(hits)}/{len(docs)}")
for name, c in sorted(hits, key=lambda x: -x[1])[:15]:
    print(f"  {c:<5} {name}")
