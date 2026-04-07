"""Print paragraph properties (style, indent, spacing) for a docx."""
import zipfile
import xml.etree.ElementTree as ET
import sys

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
WVAL = "{" + NS["w"] + "}val"

with zipfile.ZipFile(sys.argv[1]) as z:
    x = z.read("word/document.xml").decode("utf-8")
root = ET.fromstring(x)
body = root.find("w:body", NS)

for i, p in enumerate(body.findall(".//w:p", NS)):
    text = "".join(t.text or "" for t in p.findall(".//w:t", NS))
    ppr = p.find("w:pPr", NS)
    style = ""
    spacing = ""
    indent = ""
    numpr = ""
    if ppr is not None:
        ps = ppr.find("w:pStyle", NS)
        if ps is not None:
            style = ps.get(WVAL) or ""
        sp = ppr.find("w:spacing", NS)
        if sp is not None:
            attrs = []
            for k, v in sp.attrib.items():
                attrs.append(f"{k.split('}')[-1]}={v}")
            spacing = " ".join(attrs)
        ind = ppr.find("w:ind", NS)
        if ind is not None:
            attrs = []
            for k, v in ind.attrib.items():
                attrs.append(f"{k.split('}')[-1]}={v}")
            indent = " ".join(attrs)
        np = ppr.find("w:numPr", NS)
        if np is not None:
            ilvl = np.find("w:ilvl", NS)
            numid = np.find("w:numId", NS)
            numpr = f"ilvl={ilvl.get(WVAL) if ilvl is not None else '?'} numId={numid.get(WVAL) if numid is not None else '?'}"
    sys.stdout.buffer.write(
        (f"P{i} style={style!r} numpr={numpr} spacing={spacing} indent={indent}\n  text={text[:50]!r}\n").encode("utf-8", "replace")
    )
