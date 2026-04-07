"""Find docs with bold runs and check if Yu Gothic Bold lookup would help."""
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
            try:
                styles_xml = z.read("word/styles.xml").decode("utf-8")
            except KeyError:
                styles_xml = ""
        root = ET.fromstring(doc_xml)
        bold_run_count = 0
        bold_fonts = set()
        for r in root.findall(".//w:r", NS):
            rpr = r.find("w:rPr", NS)
            if rpr is None:
                continue
            b = rpr.find("w:b", NS)
            if b is None:
                continue
            v = b.get(f"{{{NS['w']}}}val")
            if v in ("0", "false"):
                continue
            bold_run_count += 1
            rfonts = rpr.find("w:rFonts", NS)
            if rfonts is not None:
                ea = rfonts.get(f"{{{NS['w']}}}eastAsia") or rfonts.get(f"{{{NS['w']}}}ascii") or ""
                if ea:
                    bold_fonts.add(ea)
        # Also check if doc has bold in styles
        sroot = ET.fromstring(styles_xml) if styles_xml else None
        bold_styles = 0
        if sroot is not None:
            for s in sroot.findall("w:style", NS):
                rpr = s.find("w:rPr", NS)
                if rpr is not None and rpr.find("w:b", NS) is not None:
                    b = rpr.find("w:b", NS)
                    v = b.get(f"{{{NS['w']}}}val")
                    if v not in ("0", "false"):
                        bold_styles += 1
        if bold_run_count > 0 or bold_styles > 0:
            results.append((docx.name, bold_run_count, bold_styles, bold_fonts))
    except Exception as e:
        pass

results.sort(key=lambda x: -(x[1]+x[2]))
print(f"{'docx':<55} {'bold_runs':<10} {'bold_styles':<12} fonts_in_bold_runs")
for name, br, bs, fs in results[:20]:
    print(f"{name:<55} {br:<10} {bs:<12} {sorted(fs)}")
print(f"\nTotal docs with bold: {len(results)}")
