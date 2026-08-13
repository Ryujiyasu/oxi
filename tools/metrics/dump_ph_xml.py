import sys
import zipfile
from pathlib import Path
_REPO = Path(__file__).resolve().parents[2]

sys.stdout.reconfigure(encoding="utf-8")

pptx = str(_REPO / r"pipeline_data\pptx_probes\ph_resolution.pptx")

with zipfile.ZipFile(pptx) as z:
    names = z.namelist()
    print("=== PPTX parts ===")
    for n in names:
        print(" ", n)
    # slide1 + rels
    slide = "ppt/slides/slide1.xml"
    rels = "ppt/slides/_rels/slide1.xml.rels"
    layout = "ppt/slideLayouts/slideLayout1.xml"
    if slide in names:
        print("\n=== slide1.xml ===")
        print(z.read(slide).decode("utf-8"))
    if rels in names:
        print("\n=== slide1.xml.rels ===")
        print(z.read(rels).decode("utf-8"))
    if layout in names:
        print("\n=== slideLayout1.xml ===")
        print(z.read(layout).decode("utf-8"))
