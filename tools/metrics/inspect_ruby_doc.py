"""Inspect ruby structure in a docx."""
import zipfile
import xml.etree.ElementTree as ET
import sys

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
path = sys.argv[1]

with zipfile.ZipFile(path) as z:
    x = z.read("word/document.xml").decode("utf-8")
root = ET.fromstring(x)
body = root.find("w:body", NS)

for i, p in enumerate(body.findall(".//w:p", NS)):
    text = "".join(t.text or "" for t in p.findall(".//w:t", NS))
    rubies = p.findall(".//w:ruby", NS)
    print(f"P{i}: {len(rubies)} ruby, text_len={len(text)}, text={text[:80]!r}")
    for j, r in enumerate(rubies):
        rt = r.find("w:rt", NS)
        rb = r.find("w:rubyBase", NS)
        rt_text = "".join(t.text or "" for t in rt.findall(".//w:t", NS)) if rt is not None else ""
        rb_text = "".join(t.text or "" for t in rb.findall(".//w:t", NS)) if rb is not None else ""
        # Check rPr for size
        rpr = r.find("w:rubyPr", NS)
        rb_size = ""
        WVAL = "{" + NS["w"] + "}val"
        if rpr is not None:
            hps = rpr.find("w:hps", NS)
            hps_base = rpr.find("w:hpsBaseText", NS)
            if hps is not None:
                rb_size = "hps=" + str(hps.get(WVAL))
            if hps_base is not None:
                rb_size += " hpsBase=" + str(hps_base.get(WVAL))
        print(f"  R{j}: base={rb_text!r} ruby={rt_text!r} {rb_size}")
