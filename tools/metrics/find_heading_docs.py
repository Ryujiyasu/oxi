"""Scan pipeline_data/docx and find docs whose paragraphs USE Heading styles."""
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

docs_dir = Path("pipeline_data/docx")
results = []
for docx in sorted(docs_dir.glob("*.docx")):
    try:
        with zipfile.ZipFile(docx) as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")
        root = ET.fromstring(doc_xml)
        body = root.find("w:body", NS)
        if body is None:
            continue
        heading_count = 0
        for p in body.findall(".//w:p", NS):
            ppr = p.find("w:pPr", NS)
            if ppr is None:
                continue
            ps = ppr.find("w:pStyle", NS)
            if ps is None:
                continue
            style_id = ps.get(f"{{{NS['w']}}}val", "")
            if "eading" in style_id or "見出" in style_id:
                heading_count += 1
        if heading_count > 0:
            results.append((docx.name, heading_count))
    except Exception:
        pass

results.sort(key=lambda x: -x[1])
print(f"{'docx':<55} headings")
for name, n in results[:30]:
    print(f"{name:<55} {n}")
print(f"\nTotal docs with headings: {len(results)}")
